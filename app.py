#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automação Progressão por Mérito — GUI (.ods -> resultado_progressao.xlsx)

- Varre 13..63 e 71..121 em cada planilha do .ods.
- Só gera linha quando (J/N/S) > 0 (após abs() e round(2)) E quando A ou B tiver conteúdo.
- Para cada J/N/S com valor: A_r, B_r, MES/ANO=C5, rubrica="00001", rendimento="r",
  sequência=C6, valor, justificativa=C9, documento legal=C10.
- Saída: apenas XLSX; 'valor' com 2 casas, sem "R$" (formato '#,##0.00'); coluna mais larga.
"""

import os, re, sys, platform
from typing import Optional, Any, List, Dict

# --- Dependências ---
try:
    import pandas as pd
except ModuleNotFoundError:
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _tk.Tk().withdraw()
    _mb.showerror("Dependência ausente", "Instale:\n\npip install pandas odfpy openpyxl")
    sys.exit(1)

# --- GUI ---
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.utils import get_column_letter

# --- Utilidades ---
def col_to_index(col: str) -> int:
    col = col.strip().upper(); idx = 0
    for ch in col:
        if not ('A' <= ch <= 'Z'): raise ValueError(f"Coluna inválida: {col}")
        idx = idx*26 + (ord(ch)-ord('A')+1)
    return idx-1

def get_cell(df: pd.DataFrame, col_letter: str, row_number: int) -> Any:
    r = row_number-1; c = col_to_index(col_letter)
    try: return df.iat[r, c]
    except Exception: return None

def parse_br_number(v: Any) -> Optional[float]:
    """Converte 'R$ 1.234,56' etc. em float; None se vazio/inválido."""
    if v is None: return None
    if isinstance(v, (int, float)):
        try: return float(v)
        except: return None
    s = str(v).strip()
    if s == "" or s.lower() in {"nan", "none"}: return None
    s = s.replace("R$","").replace("r$","").replace(" ","").replace("\xa0","")
    s = s.replace("%","").replace(".","").replace(",",".")
    s = re.sub(r"[^0-9\.\-]","", s)
    try: return float(s)
    except: return None

def not_blank(x: Any) -> bool:
    return x is not None and str(x).strip() != ""

def to_amount(v: Any) -> Optional[float]:
    """Normaliza: módulo + 2 casas; retorna None se o resultado for 0.00."""
    f = parse_br_number(v)
    if f is None: return None
    val = round(abs(float(f)), 2)
    return val if val > 0 else None

# --- Núcleo ---
def append_rows_for(df: pd.DataFrame, out: List[Dict[str, Any]], rownum: int):
    a = get_cell(df,"A",rownum)
    b = get_cell(df,"B",rownum)
    mes_ano = get_cell(df,"C",5)
    seq = get_cell(df,"C",6)
    just = get_cell(df,"C",9)
    doc = get_cell(df,"C",10)

    vals = {
        "J": to_amount(get_cell(df,"J",rownum)),
        "N": to_amount(get_cell(df,"N",rownum)),
        "S": to_amount(get_cell(df,"S",rownum)),
    }
    has_amount = any(v is not None for v in vals.values())
    has_identity = not_blank(a) or not_blank(b)

    if not (has_amount and has_identity):
        return  # evita linhas vazias/sem valor

    for _, num in vals.items():
        if num is not None:
            out.append({
                "A13": a,
                "B13": b,
                "MES/ANO": mes_ano,
                "rubrica": "00001",
                "rendimento": "r",
                "sequência": seq,
                "valor": num,
                "justificativa": just,
                "documento legal": doc,
            })

def process_sheet(df: pd.DataFrame) -> List[Dict[str, Any]]:
    linhas: List[Dict[str, Any]] = []
    for r in list(range(13, 64)) + list(range(71, 122)):
        append_rows_for(df, linhas, r)
    return linhas

def build_table_from_ods(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path): raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    try: xls = pd.ExcelFile(file_path, engine="odf")
    except Exception as e: raise RuntimeError("Falha ao abrir .ods. Instale: pip install odfpy\n\nDetalhe: "+str(e)) from e
    rows: List[Dict[str, Any]] = []
    for name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=name, header=None, engine="odf")
        rows.extend(process_sheet(df))
    cols = ["A13","B13","MES/ANO","rubrica","rendimento","sequência","valor","justificativa","documento legal"]
    return pd.DataFrame(rows, columns=cols)

def salvar_excel(df: pd.DataFrame, origem: str) -> str:
    base = os.path.dirname(os.path.abspath(origem))
    xlsx_out = os.path.join(base, "resultado_progressao.xlsx")
    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
        ws = writer.sheets["resultado"]
        if "valor" in df.columns:
            cidx = df.columns.get_loc("valor") + 1
            col_letter = get_column_letter(cidx)
            for r in range(2, len(df)+2):
                cell = ws[f"{col_letter}{r}"]
                cell.number_format = "#,##0.00"   # 1.234,56 (sem R$)
                try: cell.style = "Normal"
                except Exception: pass
            ws.column_dimensions[col_letter].width = 14  # até 000.000,00
        # Ajuste simples das demais colunas
        for i, name in enumerate(df.columns, start=1):
            if name == "valor": continue
            try:
                maxlen = max((len(str(x)) for x in [name] + df[name].astype(str).tolist()), default=10)
            except Exception:
                maxlen = 15
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, maxlen+2), 60)
    return xlsx_out

def abrir_pasta(path: str):
    """ABRE a pasta no SO (Windows/Mac/Linux)."""
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            import subprocess; subprocess.Popen(["open", path])
        else:
            import subprocess; subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

# --- GUI ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação Progressão por Mérito")
        self.geometry("560x260"); self.resizable(False, False)
        self.ods_path: Optional[str] = None

        tk.Label(self, text="Selecione um arquivo .ods para processar:", anchor="w").pack(fill="x", padx=12, pady=(12,6))
        frm = tk.Frame(self); frm.pack(fill="x", padx=12)
        self.ent_path = tk.Entry(frm); self.ent_path.pack(side="left", fill="x", expand=True, padx=(0,8))
        tk.Button(frm, text="Escolher .ods", command=self.escolher_arquivo).pack(side="left")
        self.btn_process = tk.Button(self, text="Processar", command=self.processar, state="disabled", height=2)
        self.btn_process.pack(fill="x", padx=12, pady=10)
        self.lbl_status = tk.Label(self, text="", anchor="w", fg="#333"); self.lbl_status.pack(fill="x", padx=12)
        self.btn_abrir_pasta = tk.Button(self, text="Abrir pasta de saída", command=self.abrir_saida, state="disabled")
        self.btn_abrir_pasta.pack(padx=12, pady=10)
        tk.Label(self, text="Versão 2.8 - corrigido abrir_pasta()", anchor="e", fg="#777").pack(fill="x", padx=12, pady=(8,10))

    def escolher_arquivo(self):
        p = filedialog.askopenfilename(title="Selecione o arquivo .ods",
                                       filetypes=[("Planilhas ODS","*.ods"), ("Todos os arquivos","*.*")])
        if p:
            self.ods_path = p
            self.ent_path.delete(0, tk.END); self.ent_path.insert(0, p)
            self.btn_process.config(state="normal")
            self.lbl_status.config(text="Arquivo selecionado. Pronto para processar.")

    def processar(self):
        if not self.ods_path:
            messagebox.showwarning("Aviso", "Escolha um arquivo .ods primeiro."); return
        try:
            self.lbl_status.config(text="Processando, aguarde..."); self.update_idletasks()
            df = build_table_from_ods(self.ods_path)
            xlsx = salvar_excel(df, self.ods_path)
            msg = f"Processamento concluído.\nLinhas geradas: {len(df)}\nXLSX: {xlsx}"
            self.lbl_status.config(text=msg); self.btn_abrir_pasta.config(state="normal")
            messagebox.showinfo("Concluído", msg)
        except Exception as e:
            messagebox.showerror("Erro ao processar", f"Ocorreu um erro:\n\n{e}\n\nInstale dependências:\n pip install pandas odfpy openpyxl")
            self.lbl_status.config(text=f"Erro: {e}")

    def abrir_saida(self):
        if self.ods_path:
            abrir_pasta(os.path.dirname(os.path.abspath(self.ods_path)))

if __name__ == "__main__":
    app = App(); app.mainloop()
