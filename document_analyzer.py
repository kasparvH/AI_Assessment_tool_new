import base64
import json
import io
import os
from pathlib import Path

import pandas as pd
from anthropic import Anthropic, APIError, AuthenticationError, APITimeoutError


def extract_text_from_pdf(file_bytes: bytes) -> str:
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(file_bytes))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        if not text.strip():
            raise ValueError("PDF appears to be empty or contains only images/scanned content.")
        return text
    except ImportError:
        raise RuntimeError("pypdf is not installed. Run: pip install pypdf")


def extract_text_from_docx(file_bytes: bytes) -> str:
    try:
        from docx import Document
        doc = Document(io.BytesIO(file_bytes))
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if not text.strip():
            raise ValueError("Word document appears to be empty.")
        return text
    except ImportError:
        raise RuntimeError("python-docx is not installed. Run: pip install python-docx")


def encode_image_base64(file_bytes: bytes) -> str:
    return base64.standard_b64encode(file_bytes).decode("utf-8")


SYSTEM_PROMPT = """
You are an AI readiness assessment analyst. Your job is to compare the content of
an uploaded document against the answers a respondent has given in an assessment.

Identify SPECIFIC inconsistencies where:
- The document clearly contradicts a selected answer (e.g., answer says high maturity,
  document says capability doesn't exist)
- The document provides evidence that suggests the selected answer may be too optimistic
  or too pessimistic

Rules:
- Only flag genuine, clear contradictions — not minor differences in wording
- Cite the specific text from the document that creates the inconsistency
- Explain WHY it is an inconsistency in plain English
- Rate severity: 'high' (direct contradiction), 'medium' (likely mismatch),
  'low' (worth noting but uncertain)
- If no inconsistencies found, return an empty list
- Return ONLY valid JSON — no markdown, no preamble

Output format (JSON array):
[
  {
    "question_id": <int>,
    "question_text": "<text>",
    "selected_answer": "<option text>",
    "document_evidence": "<exact quote or paraphrase from document>",
    "explanation": "<clear explanation of the inconsistency>",
    "severity": "high|medium|low"
  }
]
"""


def analyze_document_against_answers(
    file_name: str,
    file_bytes: bytes,
    file_type: str,
    current_answers: dict,
    questions_df: pd.DataFrame,
    already_answered_ids: list,
) -> list:
    # Check API key upfront with a clear message
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY is not set in your .env file. "
            "Add 'ANTHROPIC_API_KEY=sk-ant-...' to the .env file and restart."
        )

    client = Anthropic(api_key=api_key)

    # Build answered questions context
    # Fix: resolve both int and str keys since JSON roundtrip converts int keys to str
    def get_answer(qid):
        return current_answers.get(qid) or current_answers.get(str(qid)) or current_answers.get(int(qid) if str(qid).isdigit() else qid)

    answered_questions = []
    df_reset = questions_df.reset_index() if questions_df.index.name == "Question_id" else questions_df

    for qid in list(already_answered_ids)[:50]:
        row_match = df_reset[df_reset["Question_id"] == qid]
        if row_match.empty:
            continue
        row = row_match.iloc[0]
        opt_num = get_answer(qid)
        if opt_num not in (1, 2, 3):
            continue
        opt_text = row.get(f"Answer_options_{opt_num}", "")
        answered_questions.append({
            "id": int(qid),
            "question": str(row["Question_text"]),
            "dimension": str(row["Dimension"]),
            "selected_answer_number": opt_num,
            "selected_answer_text": str(opt_text),
        })

    if not answered_questions:
        raise ValueError(
            "No answered questions found to compare against. "
            "Answer at least a few questions before analyzing documents."
        )

    # Build message content
    if file_type == "image":
        if not file_bytes:
            raise ValueError("Image file is empty.")
        image_b64 = encode_image_base64(file_bytes)
        media_type = "image/png" if file_name.lower().endswith(".png") else "image/jpeg"
        content = [
            {
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": image_b64},
            },
            {
                "type": "text",
                "text": (
                    f"Document filename: {file_name}\n\n"
                    f"Assessment answers to check against:\n{json.dumps(answered_questions, indent=2)}\n\n"
                    "Analyze this image for inconsistencies with the answers above. Return JSON only."
                ),
            },
        ]
    else:
        if file_type == "pdf":
            doc_text = extract_text_from_pdf(file_bytes)
        elif file_type == "docx":
            doc_text = extract_text_from_docx(file_bytes)
        else:
            doc_text = file_bytes.decode("utf-8", errors="ignore")
            if not doc_text.strip():
                raise ValueError("Text file appears to be empty.")

        doc_text = doc_text[:6000] + ("...[truncated]" if len(doc_text) > 6000 else "")
        content = [{
            "type": "text",
            "text": (
                f"Document filename: {file_name}\n\nDOCUMENT CONTENT:\n{doc_text}\n\n"
                f"Assessment answers to check against:\n{json.dumps(answered_questions, indent=2)}\n\n"
                "Return JSON only."
            ),
        }]

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=2000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content}],
        )
    except AuthenticationError:
        raise RuntimeError(
            "API key is invalid or expired. Check your ANTHROPIC_API_KEY in the .env file."
        )
    except APITimeoutError:
        raise RuntimeError(
            "The analysis timed out. Please try again with a smaller document."
        )
    except APIError as e:
        raise RuntimeError(f"API error during analysis: {str(e)}")
    except Exception as e:
        raise RuntimeError(f"Unexpected error during analysis: {type(e).__name__}: {str(e)}")

    raw = response.content[0].text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    try:
        result = json.loads(raw)
        if not isinstance(result, list):
            return []
        return result
    except json.JSONDecodeError as e:
        raise RuntimeError(
            f"Could not parse the analysis result (invalid JSON). "
            f"Raw response started with: {raw[:100]}"
        )


def get_file_type(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    return {
        ".pdf": "pdf",
        ".docx": "docx",
        ".doc": "docx",
        ".png": "image",
        ".jpg": "image",
        ".jpeg": "image",
        ".txt": "txt",
    }.get(ext, "txt")

