import streamlit as st
import pandas as pd
import re
import io
import xlsxwriter
from datetime import datetime

def process_data(df, teacher, subject, course, level, language):
    # ... [keep all existing code until the Excel export section] ...

    # Export to Excel with formatting
    output = io.BytesIO()
    
    # Add nan_inf_to_errors option to handle NaN/INF values
    with pd.ExcelWriter(
        output, 
        engine='xlsxwriter', 
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    ) as writer:
        df_final.to_excel(writer, sheet_name='Sheet1', startrow=6, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Create new formats
        header_format = workbook.add_format({
            'bold': True, 
            'border': 1,
            'rotation': 90,
            'shrink': True
        })
        avg_header_format = workbook.add_format({
            'bold': True,
            'border': 1,
            'rotation': 90,
            'shrink': True,
            'bg_color': '#ADD8E6'  # Light blue
        })
        avg_data_format = workbook.add_format({
            'border': 1,
            'bg_color': '#ADD8E6'
        })
        border_format = workbook.add_format({'border': 1})

        # ... [keep existing header/metadata writing logic] ...

        # Write headers with appropriate formatting
        for col_num, value in enumerate(df_final.columns):
            if value.startswith(("Promedio ", "Average ")):  # Space important to avoid false matches
                worksheet.write(6, col_num, value, avg_header_format)
            else:
                worksheet.write(6, col_num, value, header_format)

        # Apply light blue background to average data cells
        average_columns = [col for col in df_final.columns 
                         if col.startswith(("Promedio ", "Average "))]  # Space important
        
        # Convert NaN values to empty strings before writing to Excel
        df_final_filled = df_final.fillna('')
        
        for col_name in average_columns:
            col_idx = df_final.columns.get_loc(col_name)
            for row_idx in range(7, 7 + len(df_final)):
                value = df_final_filled.iloc[row_idx-7, col_idx]
                worksheet.write(row_idx, col_idx, value, avg_data_format)

        # ... [keep rest of the existing code] ...

    output.seek(0)
    return output

# ... [keep rest of the code unchanged] ...
