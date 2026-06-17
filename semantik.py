import os
import re
import time
from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv()
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

SYSTEM_PROMPT = """Du er juridisk kvalitetskontrollør af danske offentlige
udbudsdokumenter. Du får forelagt en krydshenvisning fra et dokument og
indholdet af det afsnit/bilag der henvises til.

Vurdér om henvisningen giver semantisk mening — dvs. om det referred indhold
er relevant for den kontekst hvori henvisningen optræder.

Svar KUN med et JSON-objekt uden markdown-formatering:
{
  "vurdering": "OK" | "ADVARSEL" | "FEJL",
  "forklaring": "kort forklaring på dansk, max 150 tegn",
  "forslag": "konkret forslag til rettelse hvis ADVARSEL eller FEJL, ellers null"
}

Regler:
- OK: indholdet er relevant for konteksten
- ADVARSEL: indholdet er delvist relevant eller tvetydigt
- FEJL: indholdet handler om noget helt andet end konteksten tilsiger
- Vær præcis og konkret — ingen generelle bemærkninger
"""

def byg_prompt(henvisning, kilde_kontekst, ref_indhold, ref_navn):
    return f"""HENVISNING: «{henvisning}»

KONTEKST (sætningen hvori henvisningen optræder):
{kilde_kontekst}

REFERRED INDHOLD ({ref_navn}):
{ref_indhold[:1500]}

Giver denne henvisning semantisk mening?"""


def analysér_batch(fund, dokument_indhold, max_per_minut=40):
    """
    fund: liste af (samling, status, dok, ref, ctx, forkl) fra krydstjek
    dokument_indhold: dict {doknavn: fuldt tekstindhold}

    Returnerer: dict {(dok, ref, ctx): {"vurdering", "forklaring", "forslag"}}
    """
    import json

    # Kun ugyldige og usikre er interessante for semantisk analyse
    kandidater = [
        f for f in fund
        if f[1] in ("ugyldig", "usikker")
        and f[1] != "stoej"
    ]

    print(f"\nSemantisk analyse af {len(kandidater)} henvisninger...")
    resultater = {}
    interval = 60.0 / max_per_minut

    for i, (samling, status, dok, ref, ctx, forkl) in enumerate(kandidater):
        # Find referred indhold
        ref_indhold, ref_navn = _find_ref_indhold(ref, dokument_indhold)

        if not ref_indhold:
            resultater[(dok, ref, ctx)] = {
                "vurdering": "UKENDT",
                "forklaring": "Referred indhold ikke tilgængeligt for semantisk analyse",
                "forslag": None
            }
            continue

        prompt = byg_prompt(ref, ctx, ref_indhold, ref_navn)

        try:
            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=300,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": prompt}]
            )
            raw = response.content[0].text.strip()
            # Strip eventuelle markdown-backticks
            raw = re.sub(r"^```json\s*|\s*```$", "", raw, flags=re.MULTILINE)
            data = json.loads(raw)
            resultater[(dok, ref, ctx)] = data
            sym = {"OK": "✅", "ADVARSEL": "⚠️", "FEJL": "❌"}.get(
                data["vurdering"], "?")
            print(f"  {i+1}/{len(kandidater)} {sym} «{ref[:40]}» — "
                  f"{data['forklaring'][:60]}")
        except Exception as e:
            resultater[(dok, ref, ctx)] = {
                "vurdering": "FEJL",
                "forklaring": f"API-fejl: {str(e)[:80]}",
                "forslag": None
            }

        # Rate limiting
        if i < len(kandidater) - 1:
            time.sleep(interval)

    return resultater


def _find_ref_indhold(ref, dokument_indhold):
    """
    Forsøger at finde indholdet af det referred afsnit/bilag.
    Returnerer (indhold, navn) eller (None, None).
    """
    ref_lower = ref.lower()

    # Forsøg 1: bilagshenvisning — find matching dokument
    bm = re.search(
        r"\b(?:bilag|appendiks|kontraktbilag)\s+(\w+)", ref, re.IGNORECASE)
    if bm:
        nr = bm.group(1).lower().lstrip("0") or bm.group(1)
        for doknavn, indhold in dokument_indhold.items():
            dn = doknavn.lower()
            # Match fx "bilag 01", "bilag 1", "bilag 1A" mod filnavn.
            # Negativt lookahead (ikke \b) fordi filnavne ofte bruger
            # understreg lige efter nummeret (fx "Bilag 01_Tidsplan.docx"),
            # og \b matcher ikke mellem et tal og en understreg.
            if re.search(
                rf"\bbilag\s*0*{re.escape(nr)}(?![A-Za-z0-9])", dn, re.IGNORECASE):
                return indhold[:3000], doknavn

    # Forsøg 2: afsnithenvisning — find afsnit i alle dokumenter
    pm = re.search(
        r"\b(?:punkt|afsnit|underpunkt)\s+([\d.]+)", ref, re.IGNORECASE)
    if pm:
        nr = pm.group(1)
        for doknavn, indhold in dokument_indhold.items():
            # Find afsnittet i dokumentet
            m = re.search(
                rf"(?:^|\n)#{1,4}\s*{re.escape(nr)}\b(.{{0,2000}})",
                indhold, re.MULTILINE | re.DOTALL)
            if m:
                return m.group(0)[:2000], f"{doknavn} afsnit {nr}"

    return None, None
