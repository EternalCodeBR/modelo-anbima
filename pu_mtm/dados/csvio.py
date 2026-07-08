"""I/O de CSV e XLSX tolerante.

A leitura de CSV detecta o separador (',' ou ';') e aceita datas BR/ISO, porque os
CSVs do projeto passam por limpeza no Excel pt-BR. A escrita é canônica: vírgula +
datas ISO (YYYY-MM-DD).

Para arquivos .xlsx, ler_dict lê a aba "Ativos" e normaliza a coluna Familia
(nomes amigáveis → códigos internos que o motor espera).
"""
import csv
from datetime import date, datetime


# Normalização da coluna Familia: nomes amigáveis (dropdown xlsx) → códigos internos
_FAMILIA_NORM: dict[str, str] = {
    "di":           "di_puro",
    "di + spread":  "di_spread",
    "di+spread":    "di_spread",
    "prefixado":    "prefixado",
    "ipca + spread":"ipca_spread",
    "ipca+spread":  "ipca_spread",
    "ipca":         "ipca_spread",
    # também aceita os códigos internos já normalizados (CSV legado)
    "di_puro":      "di_puro",
    "di_spread":    "di_spread",
    "ipca_spread":  "ipca_spread",
}


def _cell_str(v) -> str:
    """Converte valor de célula openpyxl para string compatível com os parsers."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v == int(v):
        return str(int(v))   # 1000.0 → "1000"
    return str(v).strip()


def _ler_xlsx(caminho: str, sheet: str = "Ativos") -> list[dict]:
    """Lê aba `sheet` de um xlsx como lista de dicts; normaliza a coluna Familia."""
    import openpyxl
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    result = []
    for row in rows:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        d: dict[str, str] = {}
        for k, v in zip(header, row):
            if not k:
                continue
            s = _cell_str(v)
            if k == "Familia" and s:
                s = _FAMILIA_NORM.get(s.strip().lower(), s.strip().lower())
            d[k] = s
        result.append(d)
    wb.close()
    return result


def detectar_separador(caminho: str) -> str:
    with open(caminho, encoding="utf-8-sig") as f:
        primeira = f.readline()
    return ";" if primeira.count(";") > primeira.count(",") else ","


def ler_dict(caminho: str) -> list[dict]:
    """Lê CSV ou XLSX (aba 'Ativos') como lista de dicts."""
    if str(caminho).lower().endswith(".xlsx"):
        return _ler_xlsx(caminho)
    sep = detectar_separador(caminho)
    with open(caminho, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f, delimiter=sep))


def parse_data(s: str) -> date:
    """Aceita ISO (YYYY-MM-DD) e BR (DD/MM/AAAA)."""
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"data invalida: {s!r}")


def escrever_dict(caminho: str, campos: list[str], linhas: list[dict]) -> None:
    """Escrita canônica: vírgula, datas já em ISO nos valores."""
    with open(caminho, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(linhas)
