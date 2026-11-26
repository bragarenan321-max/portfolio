"""
Classificador de observações com IA
===================================
Lê uma planilha com observações em texto livre (por exemplo, o que o técnico
escreveu sobre um equipamento) e, para cada linha, devolve:

  - Categoria padronizada (lista fixa em categorias.json)
  - Data prometida, se o texto mencionar uma
  - Precisa de ação? (sim/não) e o motivo
  - Resumo curto
  - Confiança e origem (IA ou regras)

Com a variável ANTHROPIC_API_KEY definida, usa o Claude (API da Anthropic).
Sem a chave, roda em MODO DEMONSTRAÇÃO, com regras por palavra-chave, para
que qualquer pessoa consiga testar.

Uso:
    pip install -r requirements.txt
    python classificar.py exemplo/observacoes_tecnicos.xlsx
    python classificar.py minha_planilha.xlsx --coluna "Observação" --saida resultado.xlsx
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

AQUI = Path(__file__).parent
CATEGORIAS = json.loads((AQUI / "categorias.json").read_text(encoding="utf-8"))
MODELO = os.getenv("CLAUDE_MODEL", "claude-haiku-5-5")
LOTE = 25  # linhas por chamada à API


# ----------------------------------------------------------------------------
# Modo IA (Claude)
# ----------------------------------------------------------------------------
PROMPT_SISTEMA = """Você classifica observações escritas por técnicos de campo sobre equipamentos.
Para cada observação, responda com um objeto JSON com os campos:
  "id": o mesmo id recebido
  "categoria": exatamente uma destas opções: {categorias}
  "data_prometida": data mencionada no texto no formato DD/MM/AAAA, ou null
  "precisa_acao": true se alguém do escritório precisa agir (cobrar, abrir chamado, descontar), senão false
  "motivo_acao": frase curta explicando a ação necessária, ou null
  "resumo": no máximo 12 palavras, em português
  "confianca": número de 0 a 1
Regras: não invente informação que não está no texto; se estiver ambíguo, use a categoria "Outros"
e confiança baixa. Responda SOMENTE com uma lista JSON, sem texto antes ou depois."""


def classificar_com_ia(linhas: list[dict]) -> list[dict]:
    import anthropic  # importado só quando há chave

    cliente = anthropic.Anthropic()
    sistema = PROMPT_SISTEMA.format(categorias=", ".join(f'"{c}"' for c in CATEGORIAS))
    resultados: list[dict] = []
    for i in range(0, len(linhas), LOTE):
        lote = linhas[i:i + LOTE]
        entrada = json.dumps([{"id": l["id"], "observacao": l["texto"]} for l in lote], ensure_ascii=False)
        resp = cliente.messages.create(
            model=MODELO,
            max_tokens=4000,
            system=sistema,
            messages=[{"role": "user", "content": entrada}],
        )
        texto = "".join(b.text for b in resp.content if b.type == "text")
        bruto = json.loads(texto[texto.find("["): texto.rfind("]") + 1])
        por_id = {str(r.get("id")): r for r in bruto}
        for l in lote:
            r = por_id.get(str(l["id"]))
            if not r or r.get("categoria") not in CATEGORIAS:
                r = classificar_por_regras(l["texto"]) | {"id": l["id"]}   # resposta inválida -> regras
                r["origem"] = "regras (resposta da IA inválida)"
            else:
                r["origem"] = "IA"
            resultados.append(r)
        print(f"  {min(i + LOTE, len(linhas))}/{len(linhas)} linhas classificadas")
    return resultados


# ----------------------------------------------------------------------------
# Modo demonstração (sem chave): regras por palavra-chave
# ----------------------------------------------------------------------------
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s).lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


REGRAS = [
    ("Extraviado / perdido", r"perd|extravi|roubad|furtad|sumiu|nao (acha|encontr)"),
    ("Cliente não localizado", r"cliente (nao|ausente)|ninguem (em casa|atendeu)|endereco (nao|errado)|nao atende"),
    ("Equipamento com defeito", r"defeit|queimad|nao liga|quebrad|danific|travand"),
    ("Entregue no almoxarifado", r"(entreg|devolv)\w*\b.{0,30}\b(almox|estoque|base)"),
    ("Aguardando retirada", r"agend|retirar|retirada|vai buscar|passar (amanha|dia)|buscar (amanha|dia)"),
    ("Instalado no cliente", r"instalad|ativo no cliente|em uso|cliente (usando|com o equipamento)"),
    ("Com técnico", r"com o tecnico|na mochila|no carro|em posse|esta comigo"),
]


def classificar_por_regras(texto: str) -> dict:
    t = _norm(texto)
    categoria = next((c for c, padrao in REGRAS if re.search(padrao, t)), "Outros")
    m = re.search(r"\b(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\b", t)
    data = None
    if m:
        ano = m.group(3) or str(datetime.now().year)
        ano = "20" + ano if len(ano) == 2 else ano
        data = f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{ano}"
    acao = categoria in ("Extraviado / perdido", "Cliente não localizado", "Equipamento com defeito", "Outros")
    motivo = {
        "Extraviado / perdido": "Abrir tratativa de perda / desconto",
        "Cliente não localizado": "Contatar o cliente e reagendar",
        "Equipamento com defeito": "Enviar para análise técnica",
        "Outros": "Ler a observação e classificar manualmente",
    }.get(categoria)
    palavras = str(texto).split()
    return {
        "categoria": categoria, "data_prometida": data, "precisa_acao": acao, "motivo_acao": motivo,
        "resumo": " ".join(palavras[:12]) + ("..." if len(palavras) > 12 else ""),
        "confianca": 0.9 if categoria != "Outros" else 0.3, "origem": "regras (modo demonstração)",
    }


# ----------------------------------------------------------------------------
# Planilha de saída
# ----------------------------------------------------------------------------
def salvar(df: pd.DataFrame, caminho: Path) -> None:
    resumo = (df.groupby("Categoria").agg(Linhas=("Categoria", "size"), Precisam_de_acao=("Precisa de ação", lambda s: (s == "Sim").sum()))
                .sort_values("Linhas", ascending=False).reset_index())
    with pd.ExcelWriter(caminho, engine="xlsxwriter") as xw:
        df.to_excel(xw, sheet_name="CLASSIFICADO", index=False)
        resumo.to_excel(xw, sheet_name="RESUMO", index=False)
        wb = xw.book
        cab = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F3864"})
        sim = wb.add_format({"bg_color": "#FFF2CC", "font_color": "#9C6500", "bold": True})
        baixa = wb.add_format({"bg_color": "#F8D7DA", "font_color": "#9C0006"})
        for nome, d in (("CLASSIFICADO", df), ("RESUMO", resumo)):
            ws = xw.sheets[nome]
            for j, col in enumerate(d.columns):
                ws.write(0, j, col, cab)
                largura = min(60, max(12, int(d[col].astype(str).str.len().quantile(0.9)) + 2))
                ws.set_column(j, j, largura)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(d), len(d.columns) - 1)
        ws = xw.sheets["CLASSIFICADO"]
        j_acao = df.columns.get_loc("Precisa de ação")
        j_conf = df.columns.get_loc("Confiança")
        ws.conditional_format(1, j_acao, len(df), j_acao, {"type": "cell", "criteria": "==", "value": '"Sim"', "format": sim})
        ws.conditional_format(1, j_conf, len(df), j_conf, {"type": "cell", "criteria": "<", "value": 0.6, "format": baixa})


def main() -> int:
    ap = argparse.ArgumentParser(description="Classifica observações em texto livre com IA.")
    ap.add_argument("planilha", type=Path)
    ap.add_argument("--coluna", default="Observação", help="coluna com o texto livre")
    ap.add_argument("--id", default="ID", help="coluna identificadora (opcional)")
    ap.add_argument("--saida", type=Path, default=None)
    a = ap.parse_args()

    df = pd.read_excel(a.planilha)
    if a.coluna not in df.columns:
        print(f"Coluna '{a.coluna}' não encontrada. Colunas: {list(df.columns)}")
        return 1
    ids = df[a.id] if a.id in df.columns else pd.Series(range(1, len(df) + 1))
    linhas = [{"id": str(i), "texto": str(t)} for i, t in zip(ids, df[a.coluna].fillna(""))]

    if os.getenv("ANTHROPIC_API_KEY"):
        print(f"Classificando {len(linhas)} linhas com IA ({MODELO})...")
        res = classificar_com_ia(linhas)
    else:
        print("ANTHROPIC_API_KEY não definida: rodando em MODO DEMONSTRAÇÃO (regras por palavra-chave).")
        res = [classificar_por_regras(l["texto"]) | {"id": l["id"]} for l in linhas]

    df["Categoria"] = [r["categoria"] for r in res]
    df["Data prometida"] = [r.get("data_prometida") or "" for r in res]
    df["Precisa de ação"] = ["Sim" if r.get("precisa_acao") else "Não" for r in res]
    df["Ação sugerida"] = [r.get("motivo_acao") or "" for r in res]
    df["Resumo"] = [r.get("resumo") or "" for r in res]
    df["Confiança"] = [round(float(r.get("confianca") or 0), 2) for r in res]
    df["Origem"] = [r.get("origem", "") for r in res]

    saida = a.saida or a.planilha.with_name(a.planilha.stem + "_classificado.xlsx")
    salvar(df, saida)
    print(f"\nOK: {saida}")
    print(df["Categoria"].value_counts().to_string())
    print(f"\nPrecisam de ação: {(df['Precisa de ação'] == 'Sim').sum()} de {len(df)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
