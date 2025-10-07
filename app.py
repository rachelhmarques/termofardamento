import streamlit as st
import pandas as pd
import io, re, os, zipfile, docx, datetime, shutil
from unidecode import unidecode

# ---------- Funções auxiliares (as mesmas do Colab) ----------

def normalize(s):
    s = str(s)
    s = unidecode(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s

def detect_header_row(df_raw, max_rows=60):
    for i in range(min(max_rows, len(df_raw))):
        row_vals = [normalize(x).upper() for x in df_raw.iloc[i, :].tolist()]
        if any("ALUNO" in v for v in row_vals):
            return i
    for i in range(min(max_rows, len(df_raw))):
        row_vals = [normalize(x).upper() for x in df_raw.iloc[i, :].tolist()]
        if (any("MAT" in v for v in row_vals) and (any("CPF" in v for v in row_vals) or any(v=="RG" or " RG" in v for v in row_vals))):
            return i
    return None

def map_columns(df):
    norm_map = {normalize(c).upper().replace(".", ""): c for c in df.columns}
    def pick(*cands):
        for cand in cands:
            for k, v in norm_map.items():
                if cand in k:
                    return v
        return None
    return {
        "ALUNO": pick("ALUNO"), "MAT.SIGE": pick("MATSIGE","MAT SIGE","MATRICULA","MAT "),
        "CPF": pick("CPF"), "RG": pick("RG"), "DT.NASC.": pick("DTNASC","DATA NASC","DATA DE NASC","NASC"),
        "FILIAÇÃO": pick("FILIACAO","FILIA","MAE","PAI"),
    }

def fmt_date(val):
    if pd.isna(val) or val is None or str(val).strip()=="": return ""
    if isinstance(val, (datetime.date, datetime.datetime, pd.Timestamp)): return pd.to_datetime(val).strftime("%d/%m/%Y")
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.notna(dt): return dt.strftime("%d/%m/%Y")
    except Exception: pass
    return str(val)

def replace_in_paragraph(paragraph, mapping):
    runs = paragraph.runs
    if not runs: return
    full_text = "".join(run.text for run in runs)
    new_text = full_text
    changed = False
    for k, v in mapping.items():
        if k in new_text:
            new_text = new_text.replace(k, v if v is not None else "")
            changed = True
    if changed:
        runs[0].text = new_text
        runs[0].bold = False
        for r in runs[1:]: r.text = ""

def replace_placeholders_doc(doc, mapping):
    for p in doc.paragraphs: replace_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs: replace_in_paragraph(p, mapping)
    for section in doc.sections:
        try:
            for p in section.header.paragraphs: replace_in_paragraph(p, mapping)
        except: pass
        try:
            for p in section.footer.paragraphs: replace_in_paragraph(p, mapping)
        except: pass

def parse_sheet_name_nice(sheet_name):
    serie, turma_ano = "", ""
    m = re.match(r"\s*(\d+)[ºª]?", sheet_name, flags=re.I)
    serie = f"{m.group(1)}ª Série" if m else sheet_name
    parts = re.split(r"[_\s-]+", sheet_name.strip(), maxsplit=1)
    curso = "Técnico em " + parts[1].replace("_", " ").title() if len(parts) > 1 else sheet_name
    turma_ano = f"{serie} / {curso}"
    return turma_ano, turma_ano

def build_mapping(nome, matricula, cpf, rg, dn, filiacao, serie_ano, turma_ano):
    return {
        "[Nome Completo do Aluno(a)]": nome, "[Número de Matrícula]": matricula,
        "[Número do RG]": rg, "[Número do CPF]": cpf, "[DD/MM/AAAA]": dn,
        "[Nome Completo da Mãe e pai]": filiacao, "[Nome Completo da Mãe e Nome Completo do Pai]": filiacao,
        "[Série/Ano]": serie_ano, "[Turma e Ano do Ensino Médio]": turma_ano,
        "[Número do CPF do Aluno(a)]": cpf, "[Nome Completo da Escola]": "EEEP ÍCARO DE SOUSA MOREIRA",
        "[Numero de Matricula]": matricula, "[Numero do RG]": rg, "[Numero do CPF]": cpf,
        "[Numero do CPF do Aluno(a)]": cpf,
    }

# ---------- Interface do Streamlit ----------

st.set_page_config(layout="wide")
st.title("Gerador de Termos de Recebimento 📄")

st.markdown("""
Esta ferramenta automatiza a criação de termos de recebimento de fardamento. Siga os passos:
1.  **Faça o upload** da planilha e do modelo.
2.  **Selecione** a turma e o modo de geração.
3.  Clique em **"Gerar Documentos"** e aguarde o botão de download.
""")

# Limpa os dados do download se os arquivos forem alterados
if 'last_excel' not in st.session_state: st.session_state.last_excel = None
if 'last_template' not in st.session_state: st.session_state.last_template = None


col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. Faça o upload da sua planilha (.xlsx)", type=["xlsx"])
with col2:
    uploaded_template = st.file_uploader("2. Faça o upload do seu modelo (.docx)", type=["docx"])

# Reseta o botão de download se um novo arquivo for enviado
if uploaded_excel and uploaded_excel.id != st.session_state.last_excel:
    if 'zip_data' in st.session_state: del st.session_state.zip_data
    st.session_state.last_excel = uploaded_excel.id
if uploaded_template and uploaded_template.id != st.session_state.last_template:
    if 'zip_data' in st.session_state: del st.session_state.zip_data
    st.session_state.last_template = uploaded_template.id


if uploaded_excel and uploaded_template:
    try:
        xls = pd.ExcelFile(uploaded_excel)
        sheet_names = xls.sheet_names
        
        selected_sheet = st.selectbox("3. Selecione a Turma (Aba):", sheet_names)
        mode = st.radio("4. Escolha o modo de geração:", ("Turma inteira", "Apenas um aluno"), horizontal=True)

        selected_student_name = None
        df_for_selection = None
        cols_for_selection = None
        
        # --- LÓGICA CORRIGIDA ---
        # Mostra o seletor de aluno ANTES do botão "Gerar"
        if mode == "Apenas um aluno":
            raw = pd.read_excel(xls, sheet_name=selected_sheet, header=None)
            hdr = detect_header_row(raw)
            if hdr is not None:
                df_for_selection = pd.read_excel(xls, sheet_name=selected_sheet, header=hdr)
                cols_for_selection = map_columns(df_for_selection)
                if cols_for_selection["ALUNO"]:
                    df_for_selection = df_for_selection[df_for_selection[cols_for_selection["ALUNO"]].notna()].copy()
                    student_list = ["Selecione um aluno..."] + df_for_selection[cols_for_selection["ALUNO"]].tolist()
                    selected_student_name = st.selectbox("5. Selecione o Aluno:", student_list)
                else:
                    st.warning("Não foi possível encontrar a coluna 'ALUNO' para a seleção individual.")
            else:
                st.warning("Não foi possível detectar o cabeçalho para listar os alunos.")

        # O botão "Gerar" agora aciona o processamento
        if st.button("🚀 Gerar Documentos"):
            with st.spinner("Analisando planilha e gerando documentos... Por favor, aguarde."):
                rows_to_process = []
                df, cols = None, None

                # Processa turma inteira
                if mode == "Turma inteira":
                    raw = pd.read_excel(xls, sheet_name=selected_sheet, header=None)
                    hdr = detect_header_row(raw)
                    if hdr is None: st.error("Cabeçalho não detectado."); st.stop()
                    df = pd.read_excel(xls, sheet_name=selected_sheet, header=hdr)
                    cols = map_columns(df)
                    if not cols["ALUNO"]: st.error("Coluna 'ALUNO' não encontrada."); st.stop()
                    df = df[df[cols["ALUNO"]].notna()].copy()
                    rows_to_process = [row for _, row in df.iterrows()]
                
                # Processa apenas um aluno (usa a seleção feita ANTES do clique)
                elif mode == "Apenas um aluno":
                    if selected_student_name and selected_student_name != "Selecione um aluno...":
                        df = df_for_selection
                        cols = cols_for_selection
                        rows_to_process.append(df[df[cols["ALUNO"]] == selected_student_name].iloc[0])

                if not rows_to_process:
                    st.warning("Nenhum aluno selecionado para processar."); st.stop()

                serie_ano, turma_ano = parse_sheet_name_nice(selected_sheet)
                out_folder = f"termos_{normalize(selected_sheet)}".replace("/", "_")
                if os.path.exists(out_folder): shutil.rmtree(out_folder)
                os.makedirs(out_folder, exist_ok=True)

                for row in rows_to_process:
                    nome = str(row[cols["ALUNO"]]).strip()
                    matricula = str(row[cols["MAT.SIGE"]]).strip() if cols["MAT.SIGE"] and pd.notna(row[cols["MAT.SIGE"]]) else ""
                    cpf = str(row[cols["CPF"]]).strip() if cols["CPF"] and pd.notna(row[cols["CPF"]]) else ""
                    rg = str(row[cols["RG"]]).strip() if cols["RG"] and pd.notna(row[cols["RG"]]) else ""
                    dn = fmt_date(row[cols["DT.NASC."]]) if cols["DT.NASC."] and pd.notna(row[cols["DT.NASC."]]) else ""
                    filiacao = str(row[cols["FILIAÇÃO"]]).strip() if cols["FILIAÇÃO"] and pd.notna(row[cols["FILIAÇÃO"]]) else ""

                    doc = docx.Document(io.BytesIO(uploaded_template.getvalue()))
                    mapping = build_mapping(nome, matricula, cpf, rg, dn, filiacao, serie_ano, turma_ano)
                    replace_placeholders_doc(doc, mapping)

                    safe_nome = re.sub(r"[\/\\]+", "_", nome).strip().replace(" ", "_")
                    out_path = os.path.join(out_folder, f"Termo_{safe_nome}.docx")
                    doc.save(out_path)
                
                zip_name = f"termos_{normalize(selected_sheet)}.zip".replace("/", "_")
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname in os.listdir(out_folder):
                        zf.write(os.path.join(out_folder, fname), arcname=fname)
                
                shutil.rmtree(out_folder)
                st.session_state.zip_data = zip_buffer.getvalue()
                st.session_state.zip_filename = zip_name
                st.experimental_rerun() # Força o recarregamento para mostrar o botão de download

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar os arquivos: {e}")

if 'zip_data' in st.session_state:
     st.download_button(
         label="✅ Clique aqui para baixar o arquivo ZIP",
         data=st.session_state.zip_data,
         file_name=st.session_state.zip_filename,
         mime="application/zip",
         on_click=lambda: st.session_state.pop('zip_data', None) # Limpa o botão após o clique
     )
