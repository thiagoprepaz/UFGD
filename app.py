#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automação Progressão por Mérito — GUI (.ods -> resultado_progressao.xlsx)
- Lê todas as planilhas do .ods.
- Linhas geradas quando J13, N13 ou S13 ≠ 0 com: A13, B13, MES/ANO=C5, rubrica="00001",
  rendimento="r", sequência=C6, valor=abs(...), justificativa=C9, documento legal=C10.
- Apenas XLSX (sem CSV). 'valor' com 2 casas, sem "R$", formato Excel '#,##0.00'
  (milhar com ponto, decimal com vírgula) e coluna mais larga.
Requisitos: pip install pandas odfpy openpyxl
"""

import os, re, sys, platform
from typing import Optional, Any, List, Dict

try:
    import pandas as pd
except ModuleNotFoundError:
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _tk.Tk().withdraw()
    _mb.showerror("Dependência ausente", "Instale:\n\npip install pandas odfpy openpyxl")
    sys.exit(1)

import tkinter as tk
from tkinter import filedialog, messagebox

def col_to_index(col: str) -> int:
    col = col.strip().upper()
    idx = 0
    for ch in col:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Coluna inválida: {col}")
        idx = idx*26 + (ord(ch)-ord('A')+1)
    return idx-1

def get_cell(df: pd.DataFrame, col_letter: str, row_number: int) -> Any:
    r = row_number-1
    c = col_to_index(col_letter)
    try:
        return df.iat[r, c]
    except Exception:
        return None

def parse_br_number(value: Any) -> Optional[float]:
    """Converte valores pt-BR (ex.: 'R$ 1.234,56') para float; None se vazio."""
    if value is None: return None
    if isinstance(value, (int, float)):
        try: return float(value)
        except: return None
    s = str(value).strip()
    if s == "" or s.lower() in {"nan","none"}: return None
    s = s.replace("R$","").replace("r$","").replace(" ","").replace("\xa0","")
    s = s.replace("%","").replace(".","").replace(",",".")
    s = re.sub(r"[^0-9\.\-]","", s)
    try: return float(s)
    except: return None

def process_sheet(df: pd.DataFrame) -> List[Dict[str, Any]]:
    linhas: List[Dict[str, Any]] = []
    a13 = get_cell(df,"A",13); b13 = get_cell(df,"B",13)
    mes_ano = get_cell(df,"C",5); seq = get_cell(df,"C",6)
    just = get_cell(df,"C",9);  doc = get_cell(df,"C",10)

    # J13
    j13 = parse_br_number(get_cell(df,"J",13))
    if j13 is not None and j13 != 0:
        linhas.append({
            "A13": a13, "B13": b13, "MES/ANO": mes_ano,
            "rubrica": "00001", "rendimento": "r", "sequência": seq,
            "valor": abs(float(j13)), "justificativa": just, "documento legal": doc
        })
    # N13
    n13 = parse_br_number(get_cell(df,"N",13))
    if n13 is not None and n13 != 0:
        linhas.append({
            "A13": a13, "B13": b13, "MES/ANO": mes_ano,
            "rubrica": "00001", "rendimento": "r", "sequência": seq,
            "valor": abs(float(n13)), "justificativa": just, "documento legal": doc
        })
    # S13 (NOVO)
    s13 = parse_br_number(get_cell(df,"S",13))
    if s13 is not None and s13 != 0:
        linhas.append({
            "A13": a13, "B13": b13, "MES/ANO": mes_ano,
            "rubrica": "00001", "rendimento": "r", "sequência": seq,
            "valor": abs(float(s13)), "justificativa": just, "documento legal": doc
        })
    return linhas

def build_table_from_ods(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    try:
        xls = pd.ExcelFile(file_path, engine="odf")
    except Exception as e:
        raise RuntimeError("Falha ao abrir .ods. Instale: pip install odfpy\n\nDetalhe: "+str(e)) from e

    rows: List[Dict[str, Any]] = []
    for name in xls.sheet_names:
        df = pd.read_excel(file_path, sheet_name=name, header=None, engine="odf")
        rows.extend(process_sheet(df))

    cols = ["A13","B13","MES/ANO","rubrica","rendimento","sequência","valor","justificativa","documento legal"]
    out = pd.DataFrame(rows, columns=cols)
    if "valor" in out.columns:
        out["valor"] = pd.to_numeric(out["valor"], errors="coerce").abs().round(2)
    return out

def salvar_excel(df: pd.DataFrame, origem: str) -> str:
    from openpyxl.utils import get_column_letter

    base = os.path.dirname(os.path.abspath(origem))
    xlsx_out = os.path.join(base, "resultado_progressao.xlsx")

    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="resultado")
        ws = writer.sheets["resultado"]

        # 'valor' como número com milhar e 2 casas, sem "R$"
        if "valor" in df.columns:
            cidx = df.columns.get_loc("valor") + 1
            col_letter = get_column_letter(cidx)
            for r in range(2, len(df) + 2):
                cell = ws[f"{col_letter}{r}"]
                cell.number_format = "#,##0.00"   # exibirá 1.234,56 no PT-BR
                try: cell.style = "Normal"
                except Exception: pass
            ws.column_dimensions[col_letter].width = 14  # para até 000.000,00

        # Ajuste simples das demais colunas
        for i, name in enumerate(df.columns, start=1):
            if name == "valor":
                continue
            try:
                maxlen = max((len(str(x)) for x in [name] + df[name].astype(str).tolist()), default=10)
            except Exception:
                maxlen = 15
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, maxlen + 2), 60)

    return xlsx_out

def abrir_pasta(path: str):
    try:
        if platform.system() == "Windows":
            os.startfile(path)  # type: ignore
        elif platform.system() == "Darwin":
            __import__("subprocess").Popen(["open", path])
        else:
            __import__("subprocess").Popen(["xdg-open", path])
    except Exception:
        pass

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
        tk.Label(self, text="Versão 2.0 - GUI (inclui S13)", anchor="e", fg="#777").pack(fill="x", padx=12, pady=(8,10))

    def escolher_arquivo(self):
        p = filedialog.askopenfilename(
            title="Selecione o arquivo .ods",
            filetypes=[("Planilhas ODS","*.ods"), ("Todos os arquivos","*.*")]
        )
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
