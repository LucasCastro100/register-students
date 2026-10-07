import os
import pandas as pd
from fpdf import FPDF
from unidecode import unidecode
from dotenv import load_dotenv

current_dir = os.path.dirname(os.path.abspath(__file__))

load_dotenv(os.path.join(current_dir, ".env"))
PASS_FIXED = os.getenv("PASS_FIXED", "")
URL_LOGIN = os.getenv("URL_LOGIN", "")

DIRETORIO_DADOS = os.getenv("DIR_DATA", current_dir)
file_path = os.path.join(DIRETORIO_DADOS, "dados.xlsx")

df = pd.read_excel(file_path)
df["RA"] = df["RA"].astype(str).str.replace(r"\.0$", "", regex=True)


def desenhar_card(pdf, x, y, row):
    nome = unidecode(str(row["NOME"])).upper()
    ra = str(row["RA"])
    turma = str(row["TURMA"]).upper()

    pdf.set_draw_color(160, 160, 160)
    pdf.set_line_width(0.3)
    pdf.set_fill_color(255, 255, 255)
    pdf.rect(x, y, 94, 135, "DF")

    pdf.set_fill_color(30, 60, 120)
    pdf.rect(x, y, 94, 20, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_xy(x, y + 4)
    pdf.cell(94, 12, "ACESSO MUNDO Z", align="C")

    pdf.set_text_color(20, 20, 20)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_xy(x + 8, y + 26)
    pdf.multi_cell(78, 7, nome, align="C")

    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.set_xy(x + 8, y + 50)
    pdf.multi_cell(78, 6, f"Turma: {turma}", align="C")

    pdf.set_draw_color(210, 210, 210)
    pdf.line(x + 12, y + 62, x + 82, y + 62)

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(140, 140, 140)
    pdf.set_xy(x + 12, y + 68)
    pdf.cell(70, 5, "USUÁRIO")
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 15, 15)
    pdf.set_xy(x + 12, y + 74)
    pdf.cell(70, 8, ra)

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(140, 140, 140)
    pdf.set_xy(x + 12, y + 88)
    pdf.cell(70, 5, "SENHA")
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 15, 15)
    pdf.set_xy(x + 12, y + 94)
    pdf.cell(70, 8, PASS_FIXED)

    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(140, 140, 140)
    pdf.set_xy(x + 12, y + 108)
    pdf.cell(70, 5, "LINK DE ACESSO")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(30, 60, 120)
    pdf.set_xy(x + 12, y + 114)
    pdf.multi_cell(70, 5, URL_LOGIN)


pdf_dir = os.path.join(DIRETORIO_DADOS, "pdf")
os.makedirs(pdf_dir, exist_ok=True)

for turma, grupo in df.groupby("TURMA", sort=False):
    nome_turma = str(turma).lower().replace("º", "").strip().replace(" ", "_")
    nome_arquivo = os.path.join(pdf_dir, f"{nome_turma}.pdf")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    for pos, (_, row) in enumerate(grupo.iterrows()):
        if pos % 4 == 0:
            pdf.add_page()
        col = pos % 2
        linha = (pos // 2) % 2
        x = 10 + col * 96
        y = 12 + linha * 138
        desenhar_card(pdf, x, y, row)
    pdf.output(nome_arquivo)
    print(f"Gerado: {os.path.basename(nome_arquivo)}")
