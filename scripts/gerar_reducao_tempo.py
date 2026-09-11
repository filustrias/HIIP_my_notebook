"""
Gera `hhfa-reducao-tempo.qmd` a partir de `HHFA_reducao_tempo_nota.docx`.

A nota é densa em números (tempos, contagens de perguntas, percentagens
por área de indicadores) distribuídos por dez tabelas. Transcrevê-la à mão
seria convidar o erro de transcrição, por isso a conversão é automática.

O autor marcou a estrutura de forma consistente, o que torna a conversão
quase toda mecânica:

    estilo Heading 1   -> secção (##)
    parágrafo a negrito -> sub-rótulo (###)
    parágrafo a itálico -> ressalva metodológica

As poucas decisões que exigem julgamento estão explícitas nas constantes
abaixo, em vez de escondidas em heurísticas.

Uso:
    python scripts/gerar_reducao_tempo.py
"""

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

RAIZ = Path(__file__).resolve().parent.parent
DOCX = Path.home() / "Downloads" / "HHFA_reducao_tempo_nota.docx"
SAIDA = RAIZ / "hhfa-reducao-tempo.qmd"

# --- decisões de apresentação, por prefixo do parágrafo -------------------

# Parágrafos a negrito que NÃO são títulos: são afirmações de resultado.
DESTAQUES = {
    "Resultado do pacote": ("tip", "Resultado do pacote da secção 1"),
}

# Parágrafos a itálico que viram aviso, com o título a usar.
RESSALVAS = {
    "Os tempos são estimados": ("warning", "Estatuto dos números"),
}

# Séries paralelas de análise: o rótulo inicial, até ao primeiro ponto,
# passa a negrito para se poder percorrer a série com os olhos.
ROTULO_CORRIDO = (
    "Serviços com denominadores pequenos.",
    "Áreas de gestão e finanças.",
    "Itens de dentro dos índices de prontidão.",
    "Desenho por módulos.",
    "Três consequências.",
    "As perguntas são partilhadas entre indicadores.",
)

# Alinhamento por tabela: 'e' esquerda, 'd' direita. Uma letra por coluna.
ALINHAMENTO = {
    0: "ed",
    1: "edd",
    2: "eee",
    3: "eeee",
    4: "eee",
    5: "edddd",
    6: "eee",
    7: "ee",
    8: "ee",
    9: "ee",
}

# Linhas cujo rótulo começa assim são totais e vão a negrito.
TOTAIS = ("Subtotal", "Total")


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
        print("  (documento aberto noutro programa; li uma cópia)")
        return docx.Document(tmp)


def limpar(s):
    return re.sub(r"\s+", " ", (s or "")).strip()


def celula(s):
    return limpar(s).replace("|", "\\|")


def estilo(p):
    return p.style.name if p.style is not None else ""


def todo_negrito(p):
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)


def todo_italico(p):
    runs = [r for r in p.runs if r.text.strip()]
    return bool(runs) and all(r.italic for r in runs)


def comeca_por(texto, chaves):
    for k in chaves:
        if texto.startswith(k):
            return k
    return None


def tabela_md(t, idx):
    linhas = []
    filas = []
    for row in t.rows:
        vals, vistos = [], set()
        for c in row.cells:
            if c._tc in vistos:
                continue
            vistos.add(c._tc)
            vals.append(celula(c.text))
        filas.append(vals)
    if not filas:
        return linhas

    n = max(len(f) for f in filas)
    alin = ALINHAMENTO.get(idx, "e" * n).ljust(n, "e")

    cab = (filas[0] + [""] * n)[:n]
    linhas.append("| " + " | ".join(cab) + " |")
    linhas.append("|" + "|".join(":--" if a == "e" else "--:" for a in alin) + "|")
    for f in filas[1:]:
        f = (f + [""] * n)[:n]
        if f[0].startswith(TOTAIS):
            f = [f"**{v}**" if v else v for v in f]
        linhas.append("| " + " | ".join(f) + " |")
    linhas.append("")
    return linhas


def main():
    d = abrir(DOCX)
    out, ti = [], 0
    titulo = subtitulo = None
    corpo = []

    for child in d.element.body.iterchildren():
        tag = child.tag.split("}")[-1]

        if tag == "tbl":
            corpo.extend(tabela_md(Table(child, d), ti))
            ti += 1
            continue
        if tag != "p":
            continue

        p = Paragraph(child, d)
        texto = limpar(p.text)
        if not texto:
            continue

        if titulo is None and todo_negrito(p) and estilo(p) != "Heading 1":
            titulo = texto
            continue
        if subtitulo is None and todo_italico(p):
            subtitulo = texto
            continue

        if estilo(p) == "Heading 1":
            corpo.append(f"## {texto}\n")
            continue

        k = comeca_por(texto, RESSALVAS)
        if k and todo_italico(p):
            tipo, tit = RESSALVAS[k]
            corpo.append(f'::: {{.callout-{tipo} title="{tit}"}}')
            corpo.append(texto)
            corpo.append(":::\n")
            continue

        k = comeca_por(texto, DESTAQUES)
        if k:
            tipo, tit = DESTAQUES[k]
            corpo.append(f'::: {{.callout-{tipo} title="{tit}"}}')
            corpo.append(texto.split(":", 1)[-1].strip() if ":" in texto else texto)
            corpo.append(":::\n")
            continue

        if todo_negrito(p):
            corpo.append(f"### {texto}\n")
            continue

        k = comeca_por(texto, ROTULO_CORRIDO)
        if k:
            texto = f"**{k}**{texto[len(k):]}"
        corpo.append(texto + "\n")

    out.append("---")
    out.append(f'title: "{titulo}"')
    out.append("---\n")
    if subtitulo:
        out.append(f"*{subtitulo}*\n")
    out.extend(corpo)

    SAIDA.write_text("\n".join(out), encoding="utf-8")
    print(f"Escrito: {SAIDA}")
    print(f"  {ti} tabelas, "
          f"{sum(1 for l in out if l.startswith('## '))} secções, "
          f"{sum(1 for l in out if l.startswith('### '))} sub-secções")


if __name__ == "__main__":
    main()
