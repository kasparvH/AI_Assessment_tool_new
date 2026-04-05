import json
import os
from datetime import date
from pathlib import Path

import anthropic
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")

SYSTEM_PROMPT = """
Je bent een senior AI-strategie adviseur bij Straightable Innovatie & Strategie met diepgaande expertise 
in organisatorische AI-maturiteit, digitale transformatie en AI-governance. Je schrijft een professioneel 
AI Readiness Assessment rapport voor een Nederlandse organisatie.

Schrijfstijl:
- Professioneel en helder Nederlands op B2-niveau — toegankelijk maar inhoudelijk sterk
- Schrijf in doorlopende tekst met duidelijke alinea's; gebruik bullets alleen ter ondersteuning van de tekst, nooit als vervanging van een inhoudelijke redenering
- Wees direct, concreet en organisatiespecifiek — vermijd generieke adviestaal
- Gebruik de toon van een betrouwbare adviseur die de organisatie vooruit wil helpen
- Schrijf in actieve zinnen; vermijd passieve constructies

Gebruik exact deze secties (## koppen):

## Managementsamenvatting
Maximaal 200 woorden in doorlopende tekst. Benoem het overall maturiteitsniveau en de bijbehorende 
score, de twee of drie meest significante bevindingen, en wat dit betekent voor de AI-reis van de 
organisatie. Wees specifiek — verwijs naar daadwerkelijke dimensienamen en scores. Sluit af met 
één zin over de belangrijkste prioriteit voor de komende periode.

## Sterke punten
Behandel de drie hoogst scorende dimensies. Schrijf per dimensie een alinea van vier tot zes zinnen 
die beschrijft wat de score zegt over de huidige capaciteit, welke concrete kansen deze sterkte biedt 
en hoe de organisatie dit als strategisch voordeel kan benutten.

## Verbeterpunten
Behandel de drie laagst scorende dimensies. Schrijf per dimensie een alinea die beschrijft wat de lage 
score in de praktijk betekent, welke risico's dit met zich meebrengt als er niets verandert, en 
welke eerste concrete stap de organisatie kan zetten.

## Prioritaire aanbevelingen
Geef vijf tot zeven concrete, geprioriteerde acties. Gebruik voor elke aanbeveling het volgende formaat:

**[Korte actietitel]**
Wat te doen: [specifiek, geen algemeenheden]
Waarom dit belangrijk is: [koppel aan dimensiescores en bedrijfsimpact]
Doorlooptijd: Kort (0-3 maanden) / Middellang (3-12 maanden) / Lang (1-3 jaar)
Verwacht effect: [concreet resultaat op AI-maturiteit]

## Compliance & Governance gereedheid
Analyseer de dekkingsgraad van de vier compliance-kaders (EU AI Act, NIST AI RMF, ISO 42001, AI TRiSM).
Leg uit welke kaders een lage dekking hebben, wat de regulatoire risicos zijn, en geef een advies 
per kader. Als er inconsistenties zijn gedetecteerd uit geüploade documenten, vermeld dan dat er 
{inconsistencies_count} potentiele afwijkingen zijn vastgesteld die verificatie vereisen.

## Conclusie
Maximaal 150 woorden in doorlopende tekst. Motiverende afsluiting specifiek voor het maturiteitsniveau 
en de sector. Benadruk de weg vooruit. Eindig met een concrete uitnodiging tot actie.
"""


def generate_report(report_input: dict) -> str:
    api_key = os.getenv("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY niet gevonden. Controleer of het .env bestand "
            "in de projectmap staat met: ANTHROPIC_API_KEY=sk-ant-..."
        )
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": (
                "Schrijf het AI Readiness Assessment rapport voor deze organisatie. "
                "Gebruik uitsluitend professioneel Nederlands.\n\n"
                + json.dumps(report_input, indent=2, ensure_ascii=False)
            ),
        }],
    )
    return response.content[0].text


def _fallback_report(data: dict) -> str:
    lines = [
        "## Managementsamenvatting",
        f"{data['organisation']} heeft een totaalscore behaald van {data['overall_score_0_5']:.1f}/5.0 "
        f"({data['overall_pct']:.0f}%), wat overeenkomt met het maturiteitsniveau '{data['maturity_tier']}'. "
        "Het volledige rapport kon niet worden gegenereerd — probeer het opnieuw.",
        "",
        "## Scores per dimensie",
    ]
    for d in sorted(data["dimensions"], key=lambda x: x["score_pct"], reverse=True):
        lines.append(f"- {d['name']}: {d['score_0_5']:.1f}/5.0 ({d['score_pct']:.0f}%)")
    return "\n".join(lines)


def generate_report_safe(report_input: dict) -> tuple:
    for attempt in range(2):
        try:
            return generate_report(report_input), True
        except anthropic.APITimeoutError:
            if attempt == 1:
                return _fallback_report(report_input), False
        except anthropic.APIError as e:
            return f"Rapportgeneratie mislukt: {e}", False
    return _fallback_report(report_input), False


def build_report_input(org_name, respondent_name, sector, overall, dim_summary, compliance, inconsistencies):
    dims_sorted = dim_summary.sort_values("score_pct", ascending=False)
    dimensions = [
        {
            "name": row["dimension"],
            "score_0_5": row["score_0_5"],
            "score_pct": row["score_pct"],
            "n_questions": row["n_questions"],
            "n_answered": row["n_answered"],
            "rank": i + 1,
        }
        for i, (_, row) in enumerate(dims_sorted.iterrows())
    ]
    return {
        "organisation":          org_name,
        "respondent":            respondent_name,
        "sector":                sector,
        "date":                  date.today().isoformat(),
        "overall_score_0_5":     overall["overall_score_0_5"],
        "overall_pct":           overall["overall_pct"],
        "maturity_tier":         overall["maturity_tier"],
        "dimensions":            dimensions,
        "top3_strengths":        [d["name"] for d in dimensions[:3]],
        "top3_priorities":       [d["name"] for d in dimensions[-3:]],
        "compliance":            compliance,
        "inconsistencies_count": len(inconsistencies),
        "inconsistencies":       inconsistencies,
    }
