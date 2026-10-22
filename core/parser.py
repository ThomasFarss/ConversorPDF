"""Extração de lançamentos financeiros a partir de extratos em PDF.

Cobre os layouts dos bancos: Cresol, Sicredi, Banco do Brasil, Santander,
Mercado Pago, Nubank, Uniprime, C6, Sicoob, BTG, Itaú, Caixa e Inter.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pdfplumber

MONTHS = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
          "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
_MON = "|".join(MONTHS)

# (regex, código FEBRABAN) – a ordem importa: nomes mais específicos primeiro
BANKS: list[tuple[str, str, str]] = [
    ("Cresol", r"cresol", "133"),
    ("Sicredi", r"sicredi", "748"),
    ("Sicoob", r"sicoob|bancoob", "756"),
    ("Uniprime", r"uniprime", "084"),
    ("Mercado Pago", r"mercado\s*pago|mercadopago", "323"),
    ("Nubank", r"nubank|nu pagamentos", "260"),
    ("C6 Bank", r"c6\s*bank|banco c6", "336"),
    ("BTG Pactual", r"btg", "208"),
    ("Banco Inter", r"banco inter|inter\s*&?\s*co|\binter\b", "077"),
    ("Santander", r"santander", "033"),
    ("Caixa", r"caixa econ[oô]mica|\bcaixa\b", "104"),
    ("Itaú", r"ita[uú]", "341"),
    ("Banco do Brasil", r"banco do brasil|\bbb\.com\.br|\bbb\b", "001"),
]

# data no início da linha (ou da linha de cabeçalho de um dia)
DATE_RES = [
    re.compile(r"^\s*(\d{4})-(\d{2})-(\d{2})(?!\d)"),                              # 2025-10-01
    re.compile(r"^\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})(?!\d)"),                 # 01/10/2025
    re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2})(?![\d/,])"),                     # 01/10/25
    re.compile(r"^\s*(\d{1,2})/(\d{1,2})(?![\d/,.])"),                            # 01/10
    re.compile(rf"^\s*(\d{{1,2}})\s*(?:de\s+)?({_MON})[a-zç]*\.?(?:\s*(?:de\s+)?(\d{{4}}))?(?![\w])",
               re.IGNORECASE),                                                       # 01 OUT 2025
]
TIME_RE = re.compile(r"^\s*\d{1,2}:\d{2}(?::\d{2})?\b")

AMOUNT_RE = re.compile(
    r"(?<![\w/.,])"
    r"(?P<pre>[-+–−]?\s?)(?:R\$\s?)?(?P<pre2>[-+–−]?\s?)"
    r"(?P<num>\(?\d{1,3}(?:\.\d{3})+,\d{2}\)?|\(?\d+,\d{2}\)?)"
    r"(?P<post>\s?[-–−])?"
    r"(?:\s?(?P<flag>[CD]|CR|DB)\b\.?)?"
    r"(?![\w/,])",
    re.IGNORECASE,
)
YEAR_HINT_RE = re.compile(r"\b\d{2}/\d{2}/(20\d{2})\b|\b(?:[a-zç]{3})\.?\s+(?:de\s+)?(20\d{2})\b|\b(20\d{2})\b",
                          re.IGNORECASE)
IGNORE_WORDS = (
    "saldo anterior", "saldo do dia", "saldo final", "saldo disponível", "saldo disponivel",
    "saldo total", "saldo inicial", "saldo atual", "saldo em conta", "s a l d o", "saldo parcial",
    "saldo bloqueado", "saldo c/c", "total de entradas", "total de saídas", "total de saidas",
    "total do dia", "total ", "limite", "resumo", "rendimento bruto", "saldo da conta",
)
JUNK_LINE_RE = re.compile(
    r"p[aá]gina|\bpág\b|\d+\s*(?:/|de)\s*\d+\s*$|ouvidoria|\bsac\b|www\.|http|emitido|extrato|"
    r"saldo|total|ag[eê]ncia|\bconta\b|per[ií]odo|cnpj|central de|atendimento|lan[cç]amentos?$",
    re.IGNORECASE,
)

DEBIT_HINTS = ("enviad", "pagamento", "pgto", "compra", "débito", "debito", "saque", "tarifa", "tar ",
               "envio", "boleto", "pagto", "saída", "saida", "cobrança", "cobranca", "iof", "juros",
               "anuidade", "fatura", "transferência para", "transferencia para", "pix enviado")
CREDIT_HINTS = ("recebid", "depósito", "deposito", "crédito", "credito", "estorno", "rendimento",
                "resgate", "entrada", "reembolso", "cashback", "salário", "salario")


@dataclass
class Transaction:
    date: date
    description: str
    amount: Decimal  # negativo = débito, positivo = crédito
    source: str = ""

    @property
    def kind(self) -> str:
        return "CREDIT" if self.amount >= 0 else "DEBIT"


@dataclass
class Statement:
    transactions: list[Transaction] = field(default_factory=list)
    bank: str = ""
    bank_id: str = "000"
    account: str = "0000"


def detect_bank(text: str, hint: str = "") -> tuple[str, str]:
    """Banco pelo nome do arquivo/cabeçalho; só depois pela contagem no corpo
    (descrições como "Transferência Mercado Pago" não devem decidir)."""
    for chunk in (hint.lower(), text[:600].lower()):
        for name, pattern, code in BANKS:
            if re.search(pattern, chunk):
                return name, code
    body = text[:6000].lower()
    best = max(BANKS, key=lambda b: len(re.findall(b[1], body)))
    return (best[0], best[2]) if re.search(best[1], body) else ("", "000")


def detect_account(text: str) -> str:
    m = re.search(r"(?:conta(?:\s*corrente)?|c/c|cc)\s*[:\-]?\s*(\d[\d.\-]{3,14}\d)", text[:4000], re.IGNORECASE)
    return re.sub(r"[^\d\-]", "", m.group(1)) if m else "0000"


def _match_date(line: str, default_year: int) -> tuple[date | None, int]:
    """Devolve (data, posição final do match) ou (None, 0)."""
    for i, rx in enumerate(DATE_RES):
        m = rx.match(line)
        if not m:
            continue
        g = m.groups()
        try:
            if i == 0:
                d = date(int(g[0]), int(g[1]), int(g[2]))
            elif i == 4:
                d = date(int(g[2]) if g[2] else default_year, MONTHS[g[1].lower()], int(g[0]))
            else:
                y = g[2] if len(g) > 2 and g[2] else None
                year = default_year if not y else int(y) + (2000 if len(y) == 2 else 0)
                d = date(year, int(g[1]), int(g[0]))
        except (ValueError, KeyError):
            return None, 0
        return d, m.end()
    return None, 0


def _to_decimal(m: re.Match[str]) -> Decimal | None:
    num = m.group("num")
    negative = num.startswith("(") or "-" in (m.group("pre") or "") + (m.group("pre2") or "") \
        or "−" in (m.group("pre") or "") + (m.group("pre2") or "") \
        or "–" in (m.group("pre") or "") + (m.group("pre2") or "") \
        or bool(m.group("post"))
    txt = num.strip("()").replace(".", "").replace(",", ".")
    try:
        value = Decimal(txt)
    except InvalidOperation:
        return None
    flag = (m.group("flag") or "").upper()
    if flag in ("D", "DB"):
        negative = True
    elif flag in ("C", "CR"):
        negative = False
    return -value if negative else value


def _has_explicit_sign(m: re.Match[str]) -> bool:
    return bool(m.group("flag") or m.group("post") or m.group("num").startswith("(")
                or re.search(r"[-+–−]", (m.group("pre") or "") + (m.group("pre2") or "")))


def _infer_sign(desc: str, value: Decimal) -> Decimal:
    low = desc.lower()
    deb = any(h in low for h in DEBIT_HINTS)
    cred = any(h in low for h in CREDIT_HINTS)
    if deb and not cred:
        return -abs(value)
    return value


def _clean_desc(text: str) -> str:
    text = TIME_RE.sub("", text)
    return re.sub(r"\s+", " ", text).strip(" -–|:·")


def _is_weak(desc: str) -> bool:
    """Descrição vazia ou só números (nº de documento): o histórico real está em outra linha."""
    return desc == "Sem descrição" or not re.search(r"[A-Za-zÀ-ú]{2}", desc)


def _is_header_text(text: str) -> bool:
    low = text.lower()
    return bool(re.search(r"\b(19|20)\d{2}\b", low)) or any(re.search(b[1], low) for b in BANKS)


def _is_ignored(desc: str) -> bool:
    low = desc.lower()
    return any(w in low for w in IGNORE_WORDS)


def parse_lines(lines: list[str], source: str = "") -> list[Transaction]:
    text = "\n".join(lines)
    years = [int(a or b or c) for a, b, c in YEAR_HINT_RE.findall(text)]
    default_year = max(set(years), key=years.count) if years else datetime.now().year

    out: list[Transaction] = []
    current: date | None = None  # data de cabeçalho (Nubank, Inter, Mercado Pago agrupam por dia)
    last_idx = -10               # linha do último lançamento (para descrições quebradas)
    cont = 0
    last_above: tuple[int, str] = (-10, "")

    for n, raw in enumerate(lines):
        line = raw.strip()
        if not line:
            continue
        dt, end = _match_date(line, default_year)
        if dt:
            rest = line[end:]
            # segunda data na mesma linha (Dt. movimento / Dt. balancete)
            dt2, end2 = _match_date(rest, default_year)
            if dt2:
                rest = rest[end2:]
        else:
            rest = line

        amounts = list(AMOUNT_RE.finditer(rest))

        if dt and not amounts:
            current = dt  # cabeçalho de dia
            continue
        if dt and amounts and _is_ignored(rest):
            current = dt
            continue

        tx_date = dt or current
        if not amounts or tx_date is None:
            # possível continuação de descrição da linha anterior
            if (out and n - last_idx <= 2 and cont < 2 and not dt and not JUNK_LINE_RE.search(line)
                    and len(line) <= 90 and not amounts):
                prev = out[-1].description
                joined = line if prev == "Sem descrição" else f"{prev} {line}"
                out[-1].description = _clean_desc(joined)[:255]
                cont += 1
                last_idx = n
            elif not amounts and not dt and 3 <= len(line) <= 90 and not JUNK_LINE_RE.search(line):
                above = line  # texto solto antes do lançamento (descrição acima da linha datada)
                last_above = (n, above)
            continue
        if not dt and not re.search(r"[A-Za-zÀ-ú]{3}", rest[: amounts[0].start()]) and not out:
            continue

        chosen = amounts[0]
        value = _to_decimal(chosen)
        if value is None:
            continue
        before = rest[: chosen.start()]
        desc = _clean_desc(before)
        if not desc:
            desc = _clean_desc(rest[chosen.end():])
            desc = _clean_desc(AMOUNT_RE.sub("", desc)) or "Sem descrição"
        if _is_weak(desc) and last_above[0] == n - 1 and not _is_header_text(last_above[1]):
            # Cresol e similares: histórico na linha de cima, linha datada só com documento/valor
            desc = _clean_desc(f"{last_above[1]} {'' if desc == 'Sem descrição' else desc}")
        if _is_ignored(desc):
            continue
        if not dt and JUNK_LINE_RE.search(desc):
            continue
        if not _has_explicit_sign(chosen):
            value = _infer_sign(desc, value)
        out.append(Transaction(tx_date, desc[:255], value, source))
        last_idx, cont = n, 0
    return out


def _extract_lines(path: Path, layout: bool = False) -> list[str]:
    lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=layout) if layout else page.extract_text(x_tolerance=1.5)
            lines.extend((text or "").splitlines())
    return lines


def parse_pdf(path: str | Path) -> Statement:
    path = Path(path)
    try:
        lines = _extract_lines(path)
        txs = parse_lines(lines, path.name)
        if not txs:  # alguns PDFs só saem corretamente no modo layout
            alt = _extract_lines(path, layout=True)
            alt_txs = parse_lines(alt, path.name)
            if alt_txs:
                lines, txs = alt, alt_txs
    except Exception as e:  # noqa: BLE001
        name = type(e).__name__
        if "Password" in name:
            raise ValueError(f"{path.name}: PDF protegido por senha.") from e
        raise ValueError(f"{path.name}: não foi possível ler o PDF ({name}).") from e

    full = "\n".join(lines)
    if not txs and not full.strip():
        raise ValueError(f"{path.name}: PDF sem texto (provavelmente digitalizado/imagem).")
    bank, code = detect_bank(full, path.stem)
    return Statement(txs, bank, code, detect_account(full))
