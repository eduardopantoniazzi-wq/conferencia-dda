import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import unicodedata
from rapidfuzz.distance import JaroWinkler
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date

st.set_page_config(page_title="Conferência DDA", page_icon="🔍", layout="wide")

# ─── Helpers ──────────────────────────────────────────────────────────────────

def norm(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]", " ", s.lower())).strip()

def sim(a, b) -> float:
    na, nb = norm(a), norm(b)
    return JaroWinkler.normalized_similarity(na, nb) if na and nb else 0.0

def parse_valor(s) -> float | None:
    if isinstance(s, (int, float)):
        return float(s) if not pd.isna(s) else None
    s = re.sub(r"[R$\s]", "", str(s or ""))
    if re.search(r"\d\.\d{3},", s):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        v = float(s)
        return v if v > 0 else None
    except ValueError:
        return None

def fmt_brl(v) -> str:
    if v is None: return "—"
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_data(d) -> str:
    if not d or str(d) in ("None", "nan", ""): return "—"
    if re.match(r"\d{4}-\d{2}-\d{2}", str(d)):
        y, m, dd = str(d)[:4], str(d)[5:7], str(d)[8:10]
        return f"{dd}/{m}/{y}"
    return str(d)

def tol(v: float) -> float:
    return max(v * 0.01, 1.0)

# ─── Leitura do DDA (PDF ou Excel) ────────────────────────────────────────────

def ler_dda(uploaded) -> pd.DataFrame:
    nome_arquivo = uploaded.name.lower()

    if nome_arquivo.endswith(".pdf"):
        return _ler_dda_pdf(uploaded)
    else:
        return _ler_dda_excel(uploaded)


def _ler_dda_pdf(uploaded) -> pd.DataFrame:
    re_val  = re.compile(r"([\d]{1,3}(?:\.\d{3})*,\d{2})")
    re_data = re.compile(r"\d{2}[\/\-]\d{2}[\/\-]\d{4}")

    rows = []
    with pdfplumber.open(uploaded) as pdf:
        for page in pdf.pages:

            # Tenta extração de tabela estruturada
            for table in (page.extract_tables() or []):
                if not table or len(table) < 2:
                    continue
                # detecta índice das colunas pelo cabeçalho
                header = [norm(str(c or "")) for c in table[0]]
                bi = next((j for j, h in enumerate(header) if any(k in h for k in
                           ["beneficiario", "favorecido", "nome", "sacado", "pagador"])), -1)
                vi = next((j for j, h in enumerate(header) if "valor" in h), -1)
                di = next((j for j, h in enumerate(header) if any(k in h for k in
                           ["venc", "data"])), -1)
                si = next((j for j, h in enumerate(header) if any(k in h for k in
                           ["seu num", "seu n", "num doc", "documento"])), -1)

                # se não achou pelo cabeçalho, tenta detectar pelo conteúdo das colunas
                if bi < 0 or vi < 0:
                    bi, vi, di, si = _detectar_cols_pdf_table(table)

                for row in table[1:]:
                    if not row: continue
                    benef = str(row[bi] or "").strip() if 0 <= bi < len(row) else ""
                    val_s = str(row[vi] or "").strip() if 0 <= vi < len(row) else ""
                    dat_s = str(row[di] or "").strip() if 0 <= di < len(row) else ""
                    snum  = str(row[si] or "").strip() if 0 <= si < len(row) else ""
                    valor = parse_valor(re_val.search(val_s).group(1).replace(".", "").replace(",", ".") if re_val.search(val_s) else val_s)
                    if not benef or benef in ("None", "") or not valor:
                        continue
                    dm = re_data.search(dat_s)
                    rows.append({"beneficiario": benef, "valor": valor,
                                 "data": _fmt_iso(dm.group() if dm else ""),
                                 "seuNum": _norm_snum(snum)})

            if rows:
                break  # tabela estruturada funcionou

        if not rows:
            # Fallback: extrai palavras e agrupa por linha
            for page in pdf.pages:
                words = page.extract_words(x_tolerance=4, y_tolerance=4) or []
                linhas: dict[int, list] = {}
                for w in words:
                    k = round(w["top"] / 5) * 5
                    linhas.setdefault(k, []).append(w)

                for k in sorted(linhas):
                    ws_line = sorted(linhas[k], key=lambda w: w["x0"])
                    texts = [w["text"] for w in ws_line]
                    full  = " ".join(texts)

                    vm = re_val.search(full)
                    if not vm: continue
                    valor = parse_valor(vm.group(1).replace(".", "").replace(",", "."))
                    if not valor: continue

                    dm = re_data.search(full)
                    data = _fmt_iso(dm.group() if dm else "")

                    # beneficiário: token mais longo que não seja data nem valor
                    cands = [t for t in texts if len(t) > 4
                             and not re_val.search(t) and not re_data.search(t)
                             and not re.match(r"^\d+$", t)]
                    benef = max(cands, key=len) if cands else ""
                    if len(benef) < 4: continue

                    sn = next((t for t in texts if re.match(r"^\d{6,}$", t)), "")
                    rows.append({"beneficiario": benef, "valor": valor,
                                 "data": data, "seuNum": _norm_snum(sn)})

    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["beneficiario", "valor", "data", "seuNum"])


def _detectar_cols_pdf_table(table):
    """Detecta colunas pelo conteúdo quando o cabeçalho não tem palavras-chave."""
    re_val  = re.compile(r"\d+,\d{2}")
    re_nome = re.compile(r"[A-Za-zÀ-ÿ]{4,}")
    re_num  = re.compile(r"^\d{6,}$")

    data = table[1:]
    n = max(len(r) for r in data if r)
    sv = [0]*n; sb = [0]*n; ss = [0]*n

    for row in data:
        for j, cell in enumerate(row):
            if j >= n: break
            c = str(cell or "").strip()
            if re_val.search(c): sv[j] += 1
            if re_nome.search(c) and len(c) > 6: sb[j] += 1
            if re_num.match(c): ss[j] += 1

    vi = sv.index(max(sv)) if max(sv) > 0 else -1
    used = {vi}
    bi_cands = [(j, s) for j, s in enumerate(sb) if j not in used]
    bi = max(bi_cands, key=lambda x: x[1])[0] if bi_cands else -1
    used.add(bi)
    si_cands = [(j, s) for j, s in enumerate(ss) if j not in used]
    si = max(si_cands, key=lambda x: x[1])[0] if si_cands and max(s for _,s in si_cands) > 0 else -1
    return bi, vi, -1, si


def _ler_dda_excel(uploaded) -> pd.DataFrame:
    df = pd.read_excel(uploaded, header=None, dtype=str, engine="openpyxl")
    K_D = ["data venc","vencimento","dt venc","venc","data"]
    K_B = ["beneficiario original","beneficiario","favorecido","nome","sacado","pagador"]
    K_V = ["valor (r$)","valor r$","valor"]
    K_S = ["seu numero","seu num","seu n","num doc","documento"]

    hi = ci = bi = vi = si = -1
    for i, row in df.iterrows():
        r = [norm(str(c)) for c in row]
        d = next((j for j,c in enumerate(r) if any(k in c for k in K_D)), -1)
        b = next((j for j,c in enumerate(r) if any(k in c for k in K_B)), -1)
        v = next((j for j,c in enumerate(r) if any(k in c for k in K_V)), -1)
        s = next((j for j,c in enumerate(r) if any(k in c for k in K_S)), -1)
        if b >= 0 and v >= 0:
            hi=i; ci=d; bi=b; vi=v; si=s; break

    if hi < 0:
        st.warning("⚠️ Cabeçalho do DDA não reconhecido. Cabeçalhos encontrados: "
                   + str(list(df.iloc[0])))
        return pd.DataFrame(columns=["beneficiario","valor","data","seuNum"])

    rows = []
    for _, row in df.iloc[hi+1:].iterrows():
        v = list(row)
        benef = str(v[bi]).strip() if 0<=bi<len(v) else ""
        valor = parse_valor(v[vi]) if 0<=vi<len(v) else None
        dat_s = str(v[ci]).strip() if 0<=ci<len(v) else ""
        snum  = _norm_snum(v[si]) if 0<=si<len(v) else None
        if not benef or benef in ("nan","") or not valor: continue
        rows.append({"beneficiario": benef, "valor": valor,
                     "data": _parse_data_str(dat_s), "seuNum": snum})
    return pd.DataFrame(rows)


def _fmt_iso(ds: str) -> str | None:
    m = re.match(r"(\d{2})[\/\-](\d{2})[\/\-](\d{4})", ds)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), m.group(3)
        return f"{y}-{str(b).zfill(2)}-{str(a).zfill(2)}" if a > 12 else f"{y}-{str(a).zfill(2)}-{str(b).zfill(2)}"
    return None

def _parse_data_str(ds: str) -> str | None:
    if not ds or ds == "nan": return None
    return _fmt_iso(ds)

def _norm_snum(s) -> str | None:
    if not s or str(s) in ("nan", "", "None"): return None
    limpo = re.sub(r"[^0-9]", "", str(s)).lstrip("0")
    return limpo[:-1] if len(limpo) > 1 else (limpo or None)

# ─── Leitura do Sistema ────────────────────────────────────────────────────────

def ler_sistema(uploaded) -> pd.DataFrame:
    df = pd.read_excel(uploaded, header=None, dtype=str, engine="openpyxl")
    K_D  = ["vencimento","venc","data"]
    K_N  = ["terceiro","nome","fornecedor","sacado","pagador","beneficiario"]
    K_V  = ["vlr. nom","vlr nom","valor nominal","valor"]
    K_NF = ["nota fis","nota fiscal","nf","num nota","numero nota"]

    hi = ci = ni = vi = nfi = -1
    for i, row in df.iterrows():
        r = [norm(str(c)) for c in row]
        d  = next((j for j,c in enumerate(r) if any(k in c for k in K_D)),  -1)
        n  = next((j for j,c in enumerate(r) if any(k in c for k in K_N)),  -1)
        v  = next((j for j,c in enumerate(r) if any(k in c for k in K_V)),  -1)
        nf = next((j for j,c in enumerate(r) if any(k in c for k in K_NF)), -1)
        if n >= 0 and v >= 0:
            hi=i; ci=d; ni=n; vi=v; nfi=nf; break

    if hi < 0:
        st.warning("⚠️ Cabeçalho do Sistema não reconhecido. Cabeçalhos encontrados: "
                   + str(list(df.iloc[0])))
        return pd.DataFrame(columns=["nome","valor","data","notaFis"])

    rows = []
    for _, row in df.iloc[hi+1:].iterrows():
        v = list(row)
        nome  = str(v[ni]).strip() if 0<=ni<len(v) else ""
        valor = parse_valor(v[vi]) if 0<=vi<len(v) else None
        dat_s = str(v[ci]).strip() if 0<=ci<len(v) else ""
        nf_raw = str(v[nfi]).strip() if 0<=nfi<len(v) else ""
        if not nome or nome in ("nan","") or not valor: continue
        nf = re.sub(r"[^0-9]", "", nf_raw.split("/")[0]).lstrip("0") if nf_raw and nf_raw != "nan" else ""
        rows.append({"nome": nome, "valor": valor,
                     "data": _parse_data_str(dat_s), "notaFis": nf})
    return pd.DataFrame(rows)

# ─── Cruzamento ───────────────────────────────────────────────────────────────

def cruzar(dda: pd.DataFrame, sis: pd.DataFrame, limiar: float) -> list[dict]:
    results = []
    used = set()

    for _, d in dda.iterrows():
        best_sim, best_idx = 0.0, None
        for idx, s in sis.iterrows():
            if idx in used: continue
            s_val = sim(d["beneficiario"], s["nome"])
            if s_val < limiar: continue
            if abs(d["valor"] - s["valor"]) > tol(max(d["valor"], s["valor"])): continue
            if s_val > best_sim:
                best_sim = s_val; best_idx = idx

        if best_idx is not None:
            used.add(best_idx)
            s = sis.loc[best_idx]
            ok_nf = bool(d["seuNum"] and s["notaFis"] and d["seuNum"] == s["notaFis"])
            results.append({
                "status": "ok",
                "dda_benef": d["beneficiario"], "dda_valor": d["valor"],
                "dda_data": d["data"],          "dda_seuNum": d["seuNum"],
                "sys_nome": s["nome"],           "sys_valor": s["valor"],
                "sys_data": s["data"],           "sys_nf": s["notaFis"],
                "sim": best_sim,
                "okV": abs(d["valor"]-s["valor"]) <= tol(max(d["valor"],s["valor"])),
                "okNF": ok_nf,
                "diff": d["valor"] - s["valor"],
            })
        else:
            results.append({
                "status": "soDda",
                "dda_benef": d["beneficiario"], "dda_valor": d["valor"],
                "dda_data": d["data"],          "dda_seuNum": d["seuNum"],
                "sys_nome": None, "sys_valor": None, "sys_data": None, "sys_nf": None,
                "sim": 0.0, "okV": False, "okNF": False, "diff": None,
            })

    for idx, s in sis.iterrows():
        if idx not in used:
            results.append({
                "status": "soSis",
                "dda_benef": None, "dda_valor": None, "dda_data": None, "dda_seuNum": None,
                "sys_nome": s["nome"], "sys_valor": s["valor"],
                "sys_data": s["data"], "sys_nf": s["notaFis"],
                "sim": 0.0, "okV": False, "okNF": False, "diff": None,
            })

    return results

# ─── Export ───────────────────────────────────────────────────────────────────

def exportar(results: list[dict]) -> bytes:
    wb = Workbook(); ws = wb.active; ws.title = "Conferência DDA"
    G = PatternFill("solid", fgColor="C6EFCE")
    R = PatternFill("solid", fgColor="FFC7CE")
    Y = PatternFill("solid", fgColor="FFEB9C")
    BH = PatternFill("solid", fgColor="1F4E79")
    TH = PatternFill("solid", fgColor="1F6B75")
    GH = PatternFill("solid", fgColor="595959")
    WF = Font(color="FFFFFF", bold=True)
    C  = Alignment(horizontal="center", vertical="center")
    bd = Side(style="thin", color="CCCCCC")
    B  = Border(left=bd, right=bd, top=bd, bottom=bd)

    ws.merge_cells("A1:D1"); ws["A1"] = "📄 DDA BANCÁRIO"
    ws.merge_cells("E1:H1"); ws["E1"] = "📊 SISTEMA INTERNO"
    ws.merge_cells("I1:M1"); ws["I1"] = "RESULTADO"
    for cell, fill in [("A1",BH),("E1",TH),("I1",GH)]:
        ws[cell].fill=fill; ws[cell].font=WF; ws[cell].alignment=C

    headers = ["Data Venc.","Beneficiário","Valor (R$)","Seu Número",
               "Vencimento","Terceiro","Vlr. Nom.","Nota Fis.",
               "Similaridade","Valor ✓","NF ✓","Diferença","Status"]
    ws.append(headers)
    for cell in ws[2]:
        cell.font=Font(bold=True); cell.alignment=C; cell.border=B

    SL = {"ok":"✅ Conferido","soSis":"⚠️ Só no Sistema","soDda":"🔴 Só no DDA"}
    for r in results:
        row = [
            fmt_data(r["dda_data"]), r["dda_benef"] or "", fmt_brl(r["dda_valor"]), r["dda_seuNum"] or "",
            fmt_data(r["sys_data"]), r["sys_nome"] or "", fmt_brl(r["sys_valor"]), r["sys_nf"] or "",
            f"{r['sim']*100:.0f}%" if r["sim"] else "",
            ("✅" if r["okV"]  else "❌") if r["status"]=="ok" else "",
            ("✅" if r["okNF"] else "❌") if r["status"]=="ok" else "",
            fmt_brl(r["diff"]) if r["diff"] is not None else "",
            SL[r["status"]],
        ]
        ws.append(row)
        n = ws.max_row
        fill = G if r["status"]=="ok" else (R if r["status"]=="soDda" else Y)
        for col in range(1,14):
            c = ws.cell(n,col); c.fill=fill; c.border=B; c.alignment=Alignment(vertical="center")

    for i,w in enumerate([13,36,14,16,13,36,14,16,12,8,8,14,18],1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.row_dimensions[1].height=22; ws.row_dimensions[2].height=18
    ws.freeze_panes="A3"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ─── Interface ────────────────────────────────────────────────────────────────

st.title("🔍 Conferência DDA × Sistema Interno")
st.caption("Compara boletos do DDA bancário com o sistema — detecta possíveis fraudes")

with st.sidebar:
    st.header("⚙️ Configurações")
    limiar_pct = st.slider("Similaridade mínima de nome (%)",
                           min_value=50, max_value=100, value=75, step=5)
    limiar = limiar_pct / 100.0
    st.caption("Tolerância de valor: ±1% ou ±R$1,00")

col1, col2 = st.columns(2)
with col1:
    st.subheader("📄 DDA Bancário")
    file_dda = st.file_uploader("PDF ou Excel do DDA",
                                 type=["pdf","xlsx","xls","csv"], key="dda")
    st.caption("Colunas usadas: Beneficiário Original · Valor (R$)")

with col2:
    st.subheader("📊 Sistema Interno")
    file_sis = st.file_uploader("Relatório do sistema",
                                 type=["xlsx","xls","xlsm"], key="sis")
    st.caption("Colunas usadas: Terceiro · Vlr. Nom.")

st.divider()

if file_dda and file_sis:
    if st.button("🔍 Conferir", type="primary", use_container_width=True):

        with st.spinner("Lendo arquivos..."):
            dda = ler_dda(file_dda)
            sis = ler_sistema(file_sis)

        # Diagnóstico rápido
        with st.expander("🔎 Ver dados lidos (diagnóstico)", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.caption(f"**DDA — {len(dda)} registros**")
                st.dataframe(dda[["beneficiario","valor"]].head(15), use_container_width=True)
            with c2:
                st.caption(f"**Sistema — {len(sis)} registros**")
                st.dataframe(sis[["nome","valor"]].head(15), use_container_width=True)

        if dda.empty:
            st.error("❌ Nenhum dado extraído do DDA."); st.stop()
        if sis.empty:
            st.error("❌ Nenhum dado extraído do Sistema."); st.stop()

        with st.spinner("Cruzando..."):
            results = cruzar(dda, sis, limiar)

        n_ok    = sum(1 for r in results if r["status"]=="ok")
        n_soSis = sum(1 for r in results if r["status"]=="soSis")
        n_soDda = sum(1 for r in results if r["status"]=="soDda")

        m1,m2,m3,m4 = st.columns(4)
        m1.metric("Total", len(results))
        m2.metric("✅ Conferidos", n_ok)
        m3.metric("⚠️ Só no Sistema", n_soSis)
        m4.metric("🔴 Só no DDA", n_soDda)

        if n_soDda > 0:
            st.error(f"🚨 {n_soDda} boleto(s) no DDA sem correspondência no sistema!")

        filtro = st.radio("Filtrar:", ["Todos","✅ Conferidos","⚠️ Só no Sistema","🔴 Só no DDA"],
                          horizontal=True)
        mapa = {"Todos":None,"✅ Conferidos":"ok","⚠️ Só no Sistema":"soSis","🔴 Só no DDA":"soDda"}
        filtrado = [r for r in results if not mapa[filtro] or r["status"]==mapa[filtro]]

        SL = {"ok":"✅ Conferido","soSis":"⚠️ Só no Sistema","soDda":"🔴 Só no DDA"}
        tabela = [{
            "DDA Beneficiário": r["dda_benef"] or "—",
            "DDA Valor":        fmt_brl(r["dda_valor"]),
            "DDA Data":         fmt_data(r["dda_data"]),
            "Sys Terceiro":     r["sys_nome"] or "—",
            "Sys Vlr. Nom.":    fmt_brl(r["sys_valor"]),
            "Sys Data":         fmt_data(r["sys_data"]),
            "Similaridade":     f"{r['sim']*100:.0f}%" if r["sim"] else "—",
            "Diferença":        fmt_brl(r["diff"]) if r["diff"] is not None else "—",
            "Status":           SL[r["status"]],
        } for r in filtrado]

        def colorir(row):
            s = row["Status"]
            if "Conferido" in s: return ["background-color:#C6EFCE"]*len(row)
            if "Sistema"   in s: return ["background-color:#FFEB9C"]*len(row)
            return ["background-color:#FFC7CE"]*len(row)

        df_tab = pd.DataFrame(tabela)
        st.dataframe(df_tab.style.apply(colorir, axis=1),
                     use_container_width=True,
                     height=min(600, 80 + len(tabela)*35))

        xlsx = exportar(results)
        st.download_button("⬇️ Exportar .xlsx", data=xlsx,
                           file_name=f"conferencia_dda_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
else:
    st.info("👆 Carregue os dois arquivos para iniciar a conferência.")
