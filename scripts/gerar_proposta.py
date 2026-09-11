"""
Gera `hiip-proposal-for-action.qmd` a partir de
`HIIP Proposal for Action Angola_final.docx`.

O documento de origem é um formulário da OMS: cada secção é uma pergunta
numerada dentro de uma célula de tabela, com o enunciado do formulário a
seguir ao título. Este script desmonta essa estrutura e reconstrói um
capítulo legível, preservando o texto na íntegra.

Estrutura do docx:
    tabela 0  metadados de cabeçalho (país, orçamento, equipa, aprovações)
    tabela 1  as 13 secções do formulário, uma por linha
    tabela 2  contactos
    tabela 3  anexo 1, árvore de problema e solução
    tabela 4  anexo 2, cronograma (marcado por cor de fundo, não por texto)
    image2    anexo 3, diagrama da teoria da mudança (objeto EMF embutido)

Uso:
    python scripts/gerar_proposta.py
"""

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import docx

RAIZ = Path(__file__).resolve().parent.parent
DOCX = RAIZ.parent / "HIIP Proposal for Action Angola_final.docx"
SAIDA = RAIZ / "hiip-proposal-for-action.qmd"
IMAGEM = "hiip-theory-of-change.png"

NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def abrir(caminho):
    """Abre mesmo que o Word segure um lock exclusivo sobre o ficheiro."""
    try:
        return docx.Document(caminho)
    except PermissionError:
        tmp = Path(tempfile.gettempdir()) / f"_hiip_{caminho.name}"
        try:
            shutil.copy2(caminho, tmp)
        except PermissionError:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command",
                 f"Copy-Item -LiteralPath '{caminho}' -Destination '{tmp}' -Force"],
                check=True, capture_output=True,
            )
        print(f"  (documento aberto noutro programa; li uma cópia)")
        return docx.Document(tmp)


def limpar(s):
    """Normaliza espaços e escapa o que quebraria uma tabela pipe."""
    return re.sub(r"\s+", " ", (s or "")).strip()


def celula_md(s):
    return limpar(s).replace("|", "\\|")


def titulo_e_enunciado(par):
    """Separa o título da secção (runs a negrito) do enunciado do formulário.

    O título vem sempre a negrito; o enunciado que o segue, entre parênteses,
    vem em itálico ou sem formatação.
    """
    negrito = "".join(r.text for r in par.runs if r.bold)
    titulo = limpar(negrito).rstrip(":,")
    resto = limpar(par.text)
    if titulo and resto.startswith(titulo):
        enunciado = resto[len(titulo):].lstrip(" :")
    else:  # sem runs a negrito: corta no primeiro parêntese
        m = re.match(r"([^(]+)(\(.*)?$", resto)
        titulo = limpar(m.group(1)).rstrip(":") if m else resto
        enunciado = limpar(m.group(2)) if m and m.group(2) else ""
    return titulo, enunciado.strip("()").strip()


def paragrafos(cell, saltar_primeiro=True):
    ps = [p for p in cell.paragraphs if p.text.strip()]
    return ps[1:] if saltar_primeiro else ps


def fundo(cell):
    """Cor de fundo da célula, ou None. O cronograma é marcado assim."""
    tcPr = cell._tc.find(f"{NS}tcPr")
    if tcPr is None:
        return None
    shd = tcPr.find(f"{NS}shd")
    if shd is None:
        return None
    f = shd.get(f"{NS}fill")
    return None if f in (None, "auto", "FFFFFF") else f


def main():
    d = abrir(DOCX)
    t_meta, t_secoes, t_contacto, t_arvore, t_gantt = d.tables[:5]

    out = []
    A = out.append

    A("---")
    A('title: "The HIIP Proposal for Action"')
    A("---\n")

    # ---------------- enquadramento ------------------------------------
    A(
        "This chapter reproduces the *Proposal for Action* document that "
        "defines the HIIP engagement in Angola — the mandate everything else "
        "in this notebook works against. It is the World Health Organization "
        "(WHO) form submitted for the country, cleared by the HIIP Steering "
        "Committee on 28 April 2026.\n"
    )
    A(
        "The text is reproduced in full. What has been removed are the form's "
        "own instructions to whoever filled it in — prompts such as *\"describe "
        "the issues in the health sector\"* — which carry no information about "
        "Angola. The three annexes follow the sections, as in the original.\n"
    )

    # ---------------- metadados ----------------------------------------
    # Cada célula do cabeçalho traz um par 'Etiqueta: valor' por parágrafo,
    # por isso lemos parágrafo a parágrafo em vez de concatenar a célula.
    campos = {}
    livres = []
    for row in t_meta.rows:
        vistos = set()
        for cell in row.cells:
            if cell._tc in vistos:
                continue
            vistos.add(cell._tc)
            for p in cell.paragraphs:
                linha = limpar(p.text)
                if not linha or set(linha) <= {"_"}:
                    continue
                m = re.match(r"([^:]{2,45}):\s*(.*)$", linha)
                if m and m.group(2):
                    campos.setdefault(m.group(1).strip().lower(), m.group(2).strip())
                else:
                    livres.append(linha)

    def valor_apos(rotulo):
        """Valor escrito na linha seguinte a um rótulo solto."""
        for i, l in enumerate(livres):
            if l.lower().startswith(rotulo.lower()) and i + 1 < len(livres):
                return livres[i + 1]
        return ""

    A('::: {.callout-note title="The proposal at a glance"}\n')
    A("| | |")
    A("|:--|:--|")
    linhas_meta = [
        ("Title", campos.get("title", "")),
        ("Country", campos.get("country", "")),
        ("WHO region", campos.get("who region", "")),
        ("Estimated budget", campos.get("estimated budget", "")),
        ("WHO in-kind contribution",
         next((l.split("c.")[-1].strip() for l in livres
               if l.lower().startswith("who in kind")), "")),
        ("Team leader", campos.get("team leader", "")),
        ("Regional focal point", campos.get("ro focal", "")),
        ("Headquarters focal point", campos.get("hq focal", "")),
        ("Team members", ", ".join(
            l for l in livres
            if re.fullmatch(r"[A-Z][a-z]+ [A-Z][a-z]+", l)
            and l != valor_apos("Name of WHO Country Representative"))),
        ("WHO Representative", valor_apos("Name of WHO Country Representative")),
        ("Coordination Committee", valor_apos("Coordination Committee Date")),
        ("Steering Committee", valor_apos("Steering Committee Date")),
    ]
    for rotulo, v in linhas_meta:
        if v:
            A(f"| **{rotulo}** | {celula_md(v)} |")
    A("\n:::\n")

    # ---------------- as 13 secções ------------------------------------
    vazias = []
    for i, row in enumerate(t_secoes.rows):
        cell = row.cells[0]
        ps = [p for p in cell.paragraphs if p.text.strip()]
        if not ps:
            continue
        titulo, _ = titulo_e_enunciado(ps[0])
        A(f"## {i + 1}. {titulo}\n")

        corpo = ps[1:]
        if not corpo:
            vazias.append((i + 1, titulo))
            A(
                "::: {.callout-warning title=\"Left blank in the source document\"}\n"
                "This section carries only the form's prompt in the original "
                "document; no answer was filled in.\n:::\n"
            )
            continue

        for p in corpo:
            texto = limpar(p.text)
            if not texto:
                continue
            if p.style.name == "List Paragraph":
                A(f"- {texto}\n")
            else:
                A(texto + "\n")

    # ---------------- anexo 1: árvore de problema ----------------------
    A("## Annex 1 — Problem and solution tree\n")
    A(
        "The original renders this as a seven-column table. It is reproduced "
        "here as one block per objective, which is the same content in a form "
        "that can actually be read on a screen.\n"
    )
    cabs = [limpar(c.text) for c in t_arvore.rows[0].cells]
    for row in t_arvore.rows[1:]:
        vals = [limpar(c.text) for c in row.cells]
        objetivo = vals[5]
        rotulo = objetivo.split(".")[0] if "." in objetivo[:4] else ""
        A(f"### {rotulo + ' — ' if rotulo else ''}{vals[4]}\n")
        A("| | |")
        A("|:--|:--|")
        for k, v in zip(cabs, vals):
            if k.lower() == "outputs":
                continue
            A(f"| **{celula_md(k)}** | {celula_md(v)} |")
        A("")

    # ---------------- anexo 2: cronograma ------------------------------
    A("## Annex 2 — Deliverables, activities and timeline\n")
    A(
        "In the source document the schedule is drawn with shaded cells and no "
        "text. The months below were recovered from that shading.\n"
    )
    meses = [limpar(c.text) for c in t_gantt.rows[0].cells][1:]
    A("| Deliverable / activity | " + " | ".join(meses) + " |")
    A("|:---|" + ":-:|" * len(meses))
    for row in t_gantt.rows[1:]:
        rotulo = limpar(row.cells[0].text)
        if not rotulo:
            continue
        marcas = ["●" if fundo(c) else "" for c in row.cells[1:]]
        if not any(marcas):  # linha de título de entregável
            A(f"| **{celula_md(rotulo)}** |" + " |" * len(meses))
        else:
            A(f"| {celula_md(rotulo)} | " + " | ".join(marcas) + " |")
    A("")

    # ---------------- anexo 3: teoria da mudança -----------------------
    A("## Annex 3 — Theory of change\n")
    A(
        f"![The results chain of the proposal, from inputs through outputs and "
        f"outcomes to the goal.]({IMAGEM}){{#fig-toc fig-align=\"center\"}}\n"
    )

    # ---------------- contactos ----------------------------------------
    contacto = limpar(t_contacto.rows[0].cells[0].text)
    if contacto:
        A("## Contacts\n")
        A(contacto + "\n")

    SAIDA.write_text("\n".join(out), encoding="utf-8")
    print(f"Escrito: {SAIDA}")
    print(f"  {len(t_secoes.rows)} secções, {len(t_gantt.rows) - 1} linhas de cronograma")
    if vazias:
        print("  secções em branco no original: "
              + ", ".join(f"{n} ({t})" for n, t in vazias))


if __name__ == "__main__":
    main()
