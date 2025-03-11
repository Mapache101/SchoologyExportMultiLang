import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime

def process_data(df, teacher, subject, course, level, language):
    # Drop unwanted columns if present (if they exist in the CSV)
    columns_to_drop = [
        "Nombre de usuario",
        "Promedio General",
        "Term1 - 2024",
        "Term1 - 2024 - AUTO EVAL TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO BE_SER - Puntuación de categoría",
        "Term1 - 2024 - TO DECIDE_DECIDIR - Puntuación de categoría",
        "Term1 - 2024 - TO DO_HACER - Puntuación de categoría",
        "Term1 - 2024 - TO KNOW_SABER - Puntuación de categoría"
    ]
    df.drop(columns=columns_to_drop, inplace=True, errors='ignore')
    
    # Process columns: separate those with a grading category (coded) from general ones.
    columns_info = []  # List of dicts with info on columns having "Grading Category:"
    general_columns = []  # Columns without grading category info
    columns_to_remove = {"ID de usuario único", "ID de usuario unico"}

    for i, col in enumerate(df.columns):
        # Skip columns that are in the removal set.
        if col in columns_to_remove:
            continue
        # Exclude columns that have "(Count in Grade)" or "Category Score" in the header.
        if "(Count in Grade)" in col or "Category Score" in col:
            continue
        # Check if the column header contains "Grading Category:".
        if "Grading Category:" in col:
            # Extract the category text using a regex.
            m = re.search(r'Grading Category:\s*([^,)]+)', col)
            if m:
                category = m.group(1).strip()
            else:
                category = "Unknown"
            # Remove the parenthesized part (if any) and create a cleaner name.
            base_name = col.split('(')[0].strip()
            new_name = f"{base_name} {category}".strip()
            columns_info.append({
                'original': col,
                'new_name': new_name,
                'category': category,
                'seq_num': i
            })
        else:
            general_columns.append(col)
    
    # Reorder general columns so that columns containing names appear first.
    if language == "Español":
        name_terms = ["nombre", "apellido"]
    else:
        name_terms = ["name", "first", "last"]
    name_columns = [col for col in general_columns if any(term in col.lower() for term in name_terms)]
    other_general = [col for col in general_columns if col not in name_columns]
    general_columns_reordered = name_columns + other_general

    # Create an initial ordering: general columns then the coded columns in order of appearance.
    sorted_coded = sorted(columns_info, key=lambda x: x['seq_num'])
    new_order = general_columns_reordered + [d['original'] for d in sorted_coded]

    # Build a cleaned DataFrame and rename the coded columns.
    df_cleaned = df[new_order].copy()
    rename_dict = {d['original']: d['new_name'] for d in columns_info}
    df_cleaned.rename(columns=rename_dict, inplace=True)

    # Group coded columns by the extracted grading category.
    groups = {}
    for d in columns_info:
        groups.setdefault(d['category'], []).append(d)
    # Order the groups by the first appearance (lowest seq_num) of each category.
    group_order = sorted(groups.keys(), key=lambda cat: min(d['seq_num'] for d in groups[cat]))

    final_coded_order = []
    # For each group, sort its columns by their original order and calculate an average column.
    for cat in group_order:
        group_sorted = sorted(groups[cat], key=lambda x: x['seq_num'])
        group_names = [d['new_name'] for d in group_sorted]
        # Define the average column name based on language.
        avg_col_name = f"Promedio {cat}" if language == "Español" else f"Average {cat}"
        # Convert each column to numeric and calculate the row-wise mean.
        numeric_group = df_cleaned[group_names].apply(lambda x: pd.to_numeric(x, errors='coerce'))
        df_cleaned[avg_col_name] = numeric_group.mean(axis=1)
        # Add the group columns and then the average column.
        final_coded_order.extend(group_names)
        final_coded_order.append(avg_col_name)
    
    # Final order: general columns first, then each group (with its average column)
    final_order = general_columns_reordered + final_coded_order
    df_final = df_cleaned[final_order]

    # Replace any cell with the value "Missing" with an empty string.
    df_final.replace("Missing", "", inplace=True)

    # Export to Excel with a header section for teacher info using language-specific labels.
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Write the DataFrame starting at row 7 (startrow=6) so header info can be placed above.
        df_final.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Create cell formats.
        header_format = workbook.add_format({'bold': True, 'border': 1, 'rotation': 90, 'shrink': True})
        border_format = workbook.add_format({'border': 1})

        # Set header labels based on language.
        if language == "Español":
            teacher_label = "Docente:"
            subject_label = "Área:"
            course_label = "Curso:"
            level_label = "Nivel:"
        else:
            teacher_label = "Teacher:"
            subject_label = "Subject:"
            course_label = "Class:"
            level_label = "Level:"

        worksheet.write('A1', teacher_label, border_format)
        worksheet.write('B1', teacher, border_format)
        worksheet.write('A2', subject_label, border_format)
        worksheet.write('B2', subject, border_format)
        worksheet.write('A3', course_label, border_format)
        worksheet.write('B3', course, border_format)
        worksheet.write('A4', level_label, border_format)
        worksheet.write('B4', level, border_format)
        timestamp = datetime.now().strftime("%y-%m-%d")
        worksheet.write('A5', timestamp, border_format)

        # Write the header row for the data.
        for col_num, value in enumerate(df_final.columns):
            worksheet.write(6, col_num, value, header_format)

        # Adjust column widths.
        for idx, col_name in enumerate(df_final.columns):
            if any(term in col_name.lower() for term in name_terms):
                worksheet.set_column(idx, idx, 25)
            elif (language == "Español" and col_name.startswith("Promedio")) or (language == "English" and col_name.startswith("Average")):
                worksheet.set_column(idx, idx, 7)
            else:
                worksheet.set_column(idx, idx, 5)

        num_rows = df_final.shape[0]
        num_cols = df_final.shape[1]
        data_start_row = 6
        data_end_row = 6 + num_rows
        worksheet.conditional_format(data_start_row, 0, data_end_row, num_cols - 1, {
            'type': 'formula',
            'criteria': '=TRUE',
            'format': border_format
        })
    output.seek(0)
    return output

def main():
    st.title("Griffin's CSV to Excel")
    language = st.selectbox("Select language / Seleccione idioma", ["English", "Español"])

    # Display input fields with language-specific labels.
    if language == "Español":
        teacher = st.text_input("Escriba el nombre del docente:")
        subject = st.text_input("Escriba el área:")
        course = st.text_input("Escriba el curso:")
        level = st.text_input("Escriba el nivel:")
        uploaded_file = st.file_uploader("Subir archivo CSV", type=["csv"])
    else:
        teacher = st.text_input("Enter teacher's name:")
        subject = st.text_input("Enter subject area:")
        course = st.text_input("Enter class:")
        level = st.text_input("Enter level:")
        uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            output_excel = process_data(df, teacher, subject, course, level, language)
            if language == "Español":
                download_label = "Descargar Gradebook organizado (Excel)"
                success_msg = "Procesamiento completado!"
            else:
                download_label = "Download Organized Gradebook (Excel)"
                success_msg = "Processing completed!"
            st.download_button(
                label=download_label,
                data=output_excel,
                file_name="final_cleaned_gradebook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success(success_msg)
        except Exception as e:
            error_msg = f"Ha ocurrido un error: {e}" if language == "Español" else f"An error occurred: {e}"
            st.error(error_msg)

if __name__ == "__main__":
    main()
