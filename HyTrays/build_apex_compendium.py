#!/usr/bin/env python3
"""Build a single PDF compiling everything on HyTrays Apex.

Combines `HyTrays_Apex_concept.md` (architecture memo & scorecard) and
`HyTrays_Apex_Technical_Paper.md` (the full technical paper) into one
indexed, print-ready PDF, including all referenced figures and rendering
every LaTeX equation as a typeset image (via matplotlib mathtext).

Run:
    cd HyTrays
    python3 build_apex_compendium.py

Output:
    HyTrays/output/HyTrays_Apex_Compendium.pdf
"""

import datetime
import hashlib
import pathlib
import re

import markdown
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from weasyprint import HTML

HYTRAYS_DIR = pathlib.Path(__file__).resolve().parent
OUTPUT_DIR = HYTRAYS_DIR / "output"
EQ_DIR = OUTPUT_DIR / "pdf_equations"
EQ_DIR.mkdir(parents=True, exist_ok=True)

DOCS = [
    ("c", "Part I", HYTRAYS_DIR / "HyTrays_Apex_concept.md"),
    ("t", "Part II", HYTRAYS_DIR / "HyTrays_Apex_Technical_Paper.md"),
]


# ---------------------------------------------------------------------------
# LaTeX -> image rendering (matplotlib mathtext)
# ---------------------------------------------------------------------------

def _fix_latex(src):
    """Patch a few LaTeX constructs mathtext doesn't support."""
    src = src.replace(r"\big(", r"\left(").replace(r"\big)", r"\right)")
    src = src.replace(r"\Big(", r"\left(").replace(r"\Big)", r"\right)")

    def fix_text(m):
        inner = m.group(1).replace(" ", r"\ ")
        return r"\mathrm{" + inner + "}"

    return re.sub(r"\\text\{([^}]*)\}", fix_text, src)


def render_equation(latex, fontsize=13):
    h = hashlib.sha1(latex.encode()).hexdigest()[:12]
    out = EQ_DIR / f"eq_{h}.png"
    if not out.exists():
        fig = plt.figure(figsize=(0.01, 0.01))
        fig.text(0, 0, f"${_fix_latex(latex)}$", fontsize=fontsize)
        fig.savefig(out, dpi=200, bbox_inches="tight", pad_inches=0.06, transparent=True)
        plt.close(fig)
    return out


def replace_display_math(text):
    def repl(m):
        latex = m.group(1).strip()
        img = render_equation(latex)
        rel = img.relative_to(HYTRAYS_DIR)
        return f'\n\n<div class="equation"><img src="{rel}" alt="equation"></div>\n\n'

    return re.sub(r"\$\$(.*?)\$\$", repl, text, flags=re.S)


# ---------------------------------------------------------------------------
# Markdown -> HTML, per document
# ---------------------------------------------------------------------------

def convert_doc(prefix, md_path):
    text = md_path.read_text()
    text = replace_display_math(text)

    md = markdown.Markdown(extensions=["extra", "toc", "sane_lists"])
    html = md.convert(text)

    # Namespace heading ids / internal links so the two documents can't collide.
    html = re.sub(r'id="([^"]*)"', rf'id="{prefix}-\1"', html)
    html = re.sub(r'href="#([^"]*)"', rf'href="#{prefix}-\1"', html)

    toc = []

    def walk(tokens):
        for tok in tokens:
            toc.append(
                {"level": tok["level"], "id": f'{prefix}-{tok["id"]}', "name": tok["name"]}
            )
            walk(tok.get("children", []))

    walk(md.toc_tokens)
    return html, toc


# ---------------------------------------------------------------------------
# Contents page
# ---------------------------------------------------------------------------

def build_contents(parts):
    out = ['<section class="contents-page"><h1>Contents</h1>']
    for label, toc in parts:
        out.append(f"<h2>{label}</h2><ul>")
        for tok in toc:
            if tok["level"] > 2:
                continue
            cls = "toc-l1" if tok["level"] == 1 else "toc-l2"
            out.append(f'<li class="{cls}"><a href="#{tok["id"]}">{tok["name"]}</a></li>')
        out.append("</ul>")
    out.append("</section>")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

CSS = """
@page {
    size: Letter;
    margin: 0.95in 0.85in;
    @bottom-center { content: counter(page); font-family: "Liberation Sans"; font-size: 9pt; color: #666; }
    @top-center { content: "HyTrays Apex -- Compendium"; font-family: "Liberation Sans"; font-size: 8pt; color: #999; }
}
@page :first {
    @bottom-center { content: none; }
    @top-center { content: none; }
}
@page cover {
    @bottom-center { content: none; }
    @top-center { content: none; }
}

body { font-family: "Liberation Serif", serif; font-size: 10.3pt; line-height: 1.45; color: #1a1a1a; }

h1, h2, h3, h4, h5, h6 { font-family: "Liberation Sans", sans-serif; color: #102a43; line-height: 1.25; }
h1 { font-size: 20pt; page-break-before: always; border-bottom: 2pt solid #1a4f8b; padding-bottom: 6pt; margin-top: 0; }
h2 { font-size: 14.5pt; page-break-before: always; border-bottom: 0.75pt solid #c8d6e5; padding-bottom: 3pt; margin-top: 0; }
h3 { font-size: 12pt; margin-top: 1.4em; }
h4 { font-size: 10.5pt; margin-top: 1.2em; }

p { margin: 0.55em 0; text-align: justify; }
ul, ol { margin: 0.4em 0; padding-left: 1.4em; }
li { margin: 0.15em 0; }

a { color: #1a4f8b; text-decoration: none; }

code { font-family: "Liberation Mono", monospace; font-size: 0.92em; background: #f1f3f5; padding: 0.5pt 2.5pt; border-radius: 2pt; }
pre { font-family: "Liberation Mono", monospace; font-size: 7.3pt; line-height: 1.35; background: #f1f3f5; border: 0.5pt solid #d8dde2; border-radius: 3pt; padding: 6pt 8pt; white-space: pre; overflow: hidden; }
pre code { background: none; padding: 0; font-size: 1em; }

table { border-collapse: collapse; width: 100%; margin: 0.8em 0; font-size: 7.3pt; }
th, td { border: 0.5pt solid #aab4be; padding: 2.5pt 4pt; text-align: left; vertical-align: top; }
th { background: #dde8f3; font-family: "Liberation Sans", sans-serif; font-weight: bold; }
tr:nth-child(even) td { background: #f7f9fb; }

img { max-width: 100%; }
.doc img { display: block; margin: 0.9em auto; }
.doc p em { font-size: 9.2pt; color: #444; }
.equation { text-align: center; margin: 0.7em 0; }
.equation img { max-height: 0.9in; }

blockquote { border-left: 3pt solid #c8d6e5; margin: 0.6em 0; padding: 0.2em 0 0.2em 0.8em; color: #444; }
hr { border: none; border-top: 0.75pt solid #c8d6e5; margin: 1.5em 0; }

/* Cover page */
.cover { page: cover; page-break-after: always; text-align: center; padding-top: 2.6in; }
.cover-title { font-family: "Liberation Sans", sans-serif; font-size: 34pt; color: #102a43; margin: 0; font-weight: bold; }
.cover-subtitle { font-family: "Liberation Sans", sans-serif; font-size: 15pt; color: #1a4f8b; margin: 0.4em 0 0; }
.cover-tagline { font-size: 11pt; color: #555; margin-top: 2.2em; }
.cover-date { font-size: 10pt; color: #888; margin-top: 0.6em; }
.cover-files { font-size: 9pt; color: #888; margin-top: 2.6em; font-family: "Liberation Mono", monospace; }

/* Contents page */
.contents-page { page-break-after: always; }
.contents-page h1 { page-break-before: avoid; }
.contents-page h2 { page-break-before: avoid; border-bottom: none; font-size: 12.5pt; margin-top: 1em; }
.contents-page ul { list-style: none; padding-left: 0; }
.contents-page li.toc-l1 { font-weight: bold; margin-top: 0.5em; }
.contents-page li.toc-l2 { padding-left: 1.4em; font-size: 9.5pt; }
"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parts = []
    bodies = []
    for prefix, label, path in DOCS:
        html, toc = convert_doc(prefix, path)
        parts.append((label, toc))
        bodies.append(f'<section class="doc" id="{prefix}-top">{html}</section>')

    today = datetime.date.today().isoformat()

    cover = f"""
    <section class="cover">
      <h1 class="cover-title">HyTrays Apex</h1>
      <p class="cover-subtitle">Integrated Cartridge-Grid Tray</p>
      <p class="cover-tagline">Concept &amp; Scorecard, and Technical Paper<br>
      Complete reference compendium</p>
      <p class="cover-date">Compiled {today}</p>
      <p class="cover-files">HyTrays_Apex_concept.md<br>HyTrays_Apex_Technical_Paper.md</p>
    </section>
    """

    contents = build_contents(parts)

    full_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CSS}</style></head>
<body>
{cover}
{contents}
{"".join(bodies)}
</body></html>"""

    out_path = OUTPUT_DIR / "HyTrays_Apex_Compendium.pdf"
    HTML(string=full_html, base_url=str(HYTRAYS_DIR) + "/").write_pdf(out_path)
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
