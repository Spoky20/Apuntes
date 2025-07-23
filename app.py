import streamlit as st
import pandas as pd
import subprocess
import os
import base64
from io import BytesIO

def main():
    st.title("Generador LaTeX desde Excel")
    
    # 1. Selección de archivo Excel
    uploaded_file = st.file_uploader("Archivo Excel", type=['xlsx', 'xls'])
    
    # 2. Selección de ecuación
    equation_type = st.selectbox(
        "Tipo de ecuación:",
        ["Energía", "Ley de Hooke", "Ecuación cuadrática", "Ley de Ohm"]
    )
    
    # 3. Opciones de salida
    st.write("Opciones de exportación:")
    col1, col2 = st.columns(2)
    export_pdf = col1.checkbox("Exportar a PDF", value=True)
    export_word = col2.checkbox("Exportar a Word", value=True)
    
    # 4. Botón de generación
    if st.button("Generar Documento"):
        if uploaded_file is None:
            st.error("¡Selecciona un archivo Excel válido!")
            return
        
        try:
            with st.spinner("Procesando..."):
                # Leer Excel
                df = pd.read_excel(uploaded_file, sheet_name=equation_type)
                st.success("Datos leídos correctamente del Excel.")
                
                # Generar LaTeX
                latex_code = generate_latex_code(df, equation_type)
                
                # Mostrar vista previa del LaTeX
                with st.expander("Ver código LaTeX generado"):
                    st.code(latex_code)
                
                # Guardar archivo temporal
                with open("temp.tex", "w", encoding="utf-8") as f:
                    f.write(latex_code)
                
                # Compilar
                if export_pdf:
                    subprocess.run(["pdflatex", "-interaction=nonstopmode", "temp.tex"])
                    st.success("PDF generado correctamente")
                    
                    # Mostrar PDF
                    with open("temp.pdf", "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
                        st.markdown(pdf_display, unsafe_allow_html=True)
                
                if export_word:
                    subprocess.run(["pandoc", "temp.tex", "-o", "documento_final.docx", "--mathml"])
                    st.success("Documento Word generado correctamente")
                    
                    # Descargar Word
                    with open("documento_final.docx", "rb") as f:
                        bytes_data = f.read()
                        st.download_button(
                            label="Descargar Word",
                            data=BytesIO(bytes_data),
                            file_name="documento_final.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                
                # Limpiar archivos temporales
                for ext in ['.aux', '.log', '.out']:
                    if os.path.exists("temp" + ext):
                        os.remove("temp" + ext)
        
        except Exception as e:
            st.error(f"Ocurrió un error:\n{str(e)}")

def generate_latex_code(df, equation_type):
    latex = r"""\documentclass[12pt]{article}
\usepackage[spanish]{babel}
\usepackage{amsmath}
\usepackage{booktabs}
\usepackage{array}
\usepackage{geometry}
\geometry{a4paper, margin=1.5cm}

\title{Reporte Automatizado}
\author{Generado con Python + Streamlit}
\date{\today}

\begin{document}
\maketitle
"""

    # Configuración especial para ecuación cuadrática
    if equation_type == "Ecuación cuadrática":
        latex += r"""
\begin{table}[h]
\centering
\caption{Resultados generados automáticamente}
\begin{tabular}{l c c c >{\centering\arraybackslash}p{5cm}}
\toprule
\textbf{Caso} & \textbf{a} & \textbf{b} & \textbf{c} & \textbf{Ecuación} \\
\midrule
"""
        for _, row in df.iterrows():
            latex += f"{row['Caso']} & ${row['Coef_a']}$ & ${row['Coef_b']}$ & ${row['Coef_c']}$ & $\\boxed{{x = \\frac{{-{row['Coef_b']} \\pm \\sqrt{{{row['Coef_b']}^2 - 4 \\cdot {row['Coef_a']} \\cdot {row['Coef_c']}}}}}{{2 \\cdot {row['Coef_a']}}}}}$ \\\\\n"
    
    else:  # Para otras ecuaciones
        latex += r"""
\begin{table}[h]
\centering
\caption{Resultados generados automáticamente}
\begin{tabular}{l c c >{\centering\arraybackslash}p{5cm}}
\toprule
\textbf{Caso} & \textbf{Parámetro 1} & \textbf{Parámetro 2} & \textbf{Ecuación} \\
\midrule
"""
        for _, row in df.iterrows():
            if equation_type == "Energía":
                latex += f"{row['Caso']} & ${row['Masa (m)']}$ & ${row['Exponente (n)']}$ & $\\boxed{{E = {row['Masa (m)']} \\cdot c^{{{row['Exponente (n)']}}}}}$ \\\\\n"
            elif equation_type == "Ley de Hooke":
                latex += f"{row['Caso']} & ${row['Constante_k']}$ & ${row['Desplazamiento_x']}$ & $\\boxed{{F = {row['Constante_k']} \\cdot {row['Desplazamiento_x']}}}$ \\\\\n"
            elif equation_type == "Ley de Ohm":
                latex += f"{row['Caso']} & ${row['Voltaje_V']}$ & ${row['Corriente_I']}$ & $\\boxed{{R = \\frac{{{row['Voltaje_V']}}}{{{row['Corriente_I']}}}}}$ \\\\\n"

    latex += r"""\bottomrule
\end{tabular}
\end{table}

\begin{itemize}
\item Documento generado automáticamente el \today.
\item Tipo de ecuación: """ + equation_type + r""".
\end{itemize}

\end{document}
"""
    return latex

if __name__ == "__main__":
    main()