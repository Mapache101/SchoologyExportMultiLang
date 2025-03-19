import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime

def process_data(df, teacher, subject, course, level, language):
    columns_to_drop = [
        "Nombre de usuario", "Username", "Promedio General", "Term1 - 2024",
        "Term1 - 2024 - AUTO EVAL TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Puntuación de categoría",
        "Term1 - 2024 - TO DO_HACER - Puntuación de categoría",
        "Term1 - 2024 - TO KNOW_SABER - Puntuación de categoría",
        "Unique User ID", "Overall", "2025", "Term1 - 2025", "Term2- 2025", "Term3 - 2025"
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')
    
    exclusion_phrases = ["(Count in Grade)", "Category Score", "Ungraded"]
    
    columns_info = []
    general_columns = []
    columns_to_remove = {"ID de usuario único", "ID de usuario unico"}

    for i, col in enumerate(df.columns):
        if col in columns_to_remove:
            continue
        if any(phrase in col for phrase in exclusion_phrases):
            continue
        if "Grading Category:" in col:
            m = re.search(r'Grading Category:\s*([^,)]+)', col)
            category = m.group(1).strip() if m else "Unknown"
            base_name = col.split('(')[0].strip()
            new_name = f"{base_name} {category}".strip()
            columns_info.append({'original': col, 'new_name': new_name, 'category': category, 'seq_num': i})
        else:
            general_columns.append(col)
    
    if language == "Español":
        name_terms = ["nombre", "apellido"]
    else:
        name_terms = ["name", "first", "last"]
    name_columns = [col for col in general_columns if any(term in col.lower() for term in name_terms)]
    other_general = [col for col in general_columns if col not in name_columns]
    general_columns_reordered = name_columns + other_general

    sorted_coded = sorted(columns_info, key=lambda x: x['seq_num'])
    new_order = general_columns_reordered + [d['original'] for d in sorted_coded]

    df_cleaned = df[new_order].copy()
    rename_dict = {d['original']: d['new_name'] for d in columns_info}
    df_cleaned.rename(columns=rename_dict, inplace=True)

    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)
    group_order = sorted(groups.keys(), key=lambda cat: min(d['seq_num'] for d in groups[cat]))

    final_coded_order = []
    average_columns = []

    for cat in group_order:
        group_sorted = sorted(groups[cat], key=lambda x: x['seq_num'])
        group_names = [d['new_name'] for d in group_sorted]
        avg_col_name = f"Promedio {cat}" if language == "Español" else f"Average {cat}"
        numeric_group = df_cleaned[group_names].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_cleaned[avg_col_name] = numeric_group.mean(axis=1)
        final_coded_order.extend(group_names)
        final_coded_order.append(avg_col_name)
        average_columns.append(avg_col_name)
    
    final_order = general_columns_reordered + final_coded_order
    df_final = df_cleaned[final_order]
    df_final.replace("Missing", "", inplace=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        header_format = workbook.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True})
        border_format = workbook.add_format({'border': 1})
        light_blue_format = workbook.add_format({'bg_color': '#D9E1F2', 'border': 1})

        if language == "Español":
            labels = ["Docente:", "Área:", "Curso:", "Nivel:"]
        else:
            labels = ["Teacher:", "Subject:", "Class:", "Level:"]

        worksheet.write('A1', labels[0], border_format)
        worksheet.write('B1', teacher, border_format)
        worksheet.write('A2', labels[1], border_format)
        worksheet.write('B2', subject, border_format)
        worksheet.write('A3', labels[2], border_format)
        worksheet.write('B3', course, border_format)
        worksheet.write('A4', labels[3], border_format)
        worksheet.write('B4', level, border_format)
        worksheet.write('A5', datetime.now().strftime("%y-%m-%d"), border_format)

        for col_num, value in enumerate(df_final.columns):
            worksheet.write(6, col_num, value, header_format)

        for idx, col_name in enumerate(df_final.columns):
            if col_name in average_columns:
                worksheet.set_column(idx, idx, 10, light_blue_format)
            elif any(term in col_name.lower() for term in name_terms):
                worksheet.set_column(idx, idx, 25)
            else:
                worksheet.set_column(idx, idx, 7)

        num_rows = df_final.shape[0]
        num_cols = df_final.shape[1]
        worksheet.conditional_format(6, 0, 6 + num_rows, num_cols - 1, {
            'type': 'formula', 'criteria': '=TRUE', 'format': border_format
        })
    output.seek(0)
    return output

def main():
    st.title("Griffin's CSV to Excel")
    language = st.selectbox("Select language / Seleccione idioma", ["English", "Español"])
    teacher = st.text_input("Enter teacher's name:") if language == "English" else st.text_input("Escriba el nombre del docente:")
    subject = st.text_input("Enter subject area:") if language == "English" else st.text_input("Escriba el área:")
    course = st.text_input("Enter class:") if language == "English" else st.text_input("Escriba el curso:")
    level = st.text_input("Enter level:") if language == "English" else st.text_input("Escriba el nivel:")
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            output_excel = process_data(df, teacher, subject, course, level, language)
            st.download_button("Download Organized Gradebook (Excel)", output_excel, "final_cleaned_gradebook.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.success("Processing completed!")
        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
