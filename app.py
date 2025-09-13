import streamlit as st
st.set_page_config(page_title="Fleet Repair Monthly", layout="wide")  # <--- add this as the first Streamlit call
import pandas as pd
import datetime
import openpyxl
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
from docx import Document
from docx.shared import Pt, RGBColor  # add RGBColor to your import
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

st.markdown(
    """
    <style>
    html, body, [class*="css"]  {
        font-size: 14px !important;
    }
    .stDataFrame, .stTable, .stMarkdown, .stTextInput, .stSelectbox, .stButton, .stForm, .stDateInput, .stNumberInput {
        font-size: 14px !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

sale_entry, sales = st.columns([5, 8])

with sale_entry:
    form_values = {}
    with st.form("sale_form", clear_on_submit=True):
        areas_df = pd.read_excel("static/areas.xlsx")
        area_list = areas_df['Area'].dropna().unique().tolist()

        vehicle_df = pd.read_excel("static/vehicles.xlsx")
        vehicle_list = vehicle_df['Vehicle'].dropna().unique().tolist()

        col1, col2, col3 = st.columns([3,4,3])

        with col1:
            if area_list:
                form_values["area"] = st.selectbox("Area", area_list, key="area_select")
            else:
                form_values["area"] = st.text_input("Area", key="area_text")

        with col2:
            if vehicle_list:
                form_values["vehicle"] = st.selectbox("Vehicle", vehicle_list, key="vehicle_select")
            else:
                form_values["vehicle"] = st.text_input("Vehicle", key="vehicle_id")

        with col3:
            first_day_prev_month = (datetime.date.today().replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
            form_values["date_of_repair"] = st.date_input("Date of Repair", key="date_of_repair", value=first_day_prev_month)

        if "repair_rows" not in st.session_state:
            st.session_state.repair_rows = 5

        for i in range(st.session_state.repair_rows):
            col1, col2 = st.columns(2)

            with col1:
                form_values[f"repair_description_{i}"] = st.text_input(f"Repair Description {i+1}", key=f"repair_description_{i}")

            with col2:
                form_values[f"cost_{i}"] = st.number_input(f"Cost {i+1}", min_value=0, key=f"cost_{i}", value=None, placeholder="")

        submit_pressed = st.form_submit_button("Submit Repair Entry")
        
        if submit_pressed:

            repairs_file_path = "static/repairs_excel.xlsx"
            if not os.path.exists(repairs_file_path):
                empty_df = pd.DataFrame(columns=["area", "vehicle", "date_of_repair", "repairs", "total_cost"])
                today_str = datetime.date.today().strftime("%d-%m-%Y")
                sheet_name=f"{today_str}"
                with pd.ExcelWriter(repairs_file_path, engine="openpyxl") as writer:
                    empty_df.to_excel(writer, index=False, sheet_name=sheet_name)

            file_path = os.path.join("static", "repairs_excel.xlsx")
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            today_str = datetime.date.today().strftime("%d-%m-%Y")
            sheet_name=f"{today_str}"

            repairs_excel_df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Format date_of_repair as "day, month, year"
            date_val = form_values.get("date_of_repair")
            if isinstance(date_val, (datetime.date, datetime.datetime)):
                formatted_date = date_val.strftime("%d, %B, %Y")
            else:
                formatted_date = str(date_val) if date_val else ""

            new_row = {
                "Area": form_values.get("area"),
                "Vehicle ID": form_values.get("vehicle"),
                "Date": form_values.get("date_of_repair").strftime("%d, %B, %Y"),
                "Description": "\n".join(
                    f"{(form_values.get(f'repair_description_{i}') or '').strip()} (ugx {(form_values.get(f'cost_{i}') or 0):,})"
                    for i in range(st.session_state.repair_rows)
                    if (form_values.get(f'repair_description_{i}') or '').strip() or (form_values.get(f'cost_{i}') is not None)
                ),
                "Total Cost (ugx)": "{:,}".format(
                    int(sum(
                        float(form_values.get(f"cost_{i}") or 0)
                        for i in range(st.session_state.repair_rows)
                    ))
                )
            }

            new_rows_df = pd.DataFrame([new_row])
            monthly_repairs_df = pd.concat([repairs_excel_df, new_rows_df], ignore_index=True)

            with pd.ExcelWriter(repairs_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                monthly_repairs_df.to_excel(writer, index=False, sheet_name=sheet_name)

            work_book = load_workbook(repairs_file_path)
            ws = work_book[sheet_name]

            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            col_max_width = {}

            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    cell.border = thin_border
                    val = str(cell.value) if cell.value is not None else ""
                    col_letter = cell.column_letter
                    col_max_width[col_letter] = max(col_max_width.get(col_letter, 0), len(val))

            # Set column widths based on max content length
            for col_letter, max_len in col_max_width.items():
                ws.column_dimensions[col_letter].width = max(15, min(max_len + 3, 60))  # reasonable min/max

            work_book.save(repairs_file_path)

            st.success("Repair entry submitted!")
            st.session_state.repair_rows = 5
            st.session_state.clear()

    col1, col2, col3, col4, col5 = st.columns([1, 1, 1, 1, 1])
    
    with col1:
        def add_row():
            st.session_state.repair_rows += 1
        st.button("Add", key=f"add_{i}", on_click=add_row)
    
    with col2:
        def remove_row():
            if st.session_state.repair_rows > 1:
                st.session_state.repair_rows -= 1
        st.button("Remove", key=f"remove_{i}", on_click=remove_row)


with sales:
    st.markdown("#### Sales Recorded")

    file_path = "static/repairs_excel.xlsx"
    os.makedirs(os.path.dirname(file_path), exist_ok=True)

    if not os.path.exists(file_path):
        empty_df = pd.DataFrame(columns=["Area", "Vehicle ID", "Date", "Description", "Total Cost (ugx)"])
        today_str = datetime.date.today().strftime("%d-%m-%Y")
        empty_df.to_excel(file_path, index=False, sheet_name=today_str)

    today_str = datetime.date.today().strftime("%d-%m-%Y")
    sheet_name=f"{today_str}"
    repairs_excel_df = pd.read_excel('static/repairs_excel.xlsx', sheet_name=sheet_name)

    edited_df = st.data_editor(
        repairs_excel_df,
        num_rows="dynamic",
        width="stretch",
        key="data_editor"
    )

    col1, col2, col3 = st.columns([1, 1, 1])

    with col1:
        if st.button("Save Changes"):
            monthly_repairs_df = edited_df.copy()
            repairs_file_path = os.path.join("static", "repairs_excel.xlsx")
            today_str = datetime.date.today().strftime("%d-%m-%Y")
            sheet_name = f"{today_str}"


            try:
                with pd.ExcelWriter(repairs_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    monthly_repairs_df.to_excel(writer, index=False, sheet_name=sheet_name)
            except FileNotFoundError:
                with pd.ExcelWriter(repairs_file_path, engine="openpyxl") as writer:
                    monthly_repairs_df.to_excel(writer, index=False, sheet_name=sheet_name)

            work_book = load_workbook(repairs_file_path)
            ws = work_book[sheet_name]

            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            col_max_width = {}

            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    cell.border = thin_border
                    val = str(cell.value) if cell.value is not None else ""
                    col_letter = cell.column_letter
                    col_max_width[col_letter] = max(col_max_width.get(col_letter, 0), len(val))

            # Set column widths based on max content length
            for col_letter, max_len in col_max_width.items():
                ws.column_dimensions[col_letter].width = max(15, min(max_len + 3, 60))  # reasonable min/max

            work_book.save(repairs_file_path)
        
    with col2:
        if st.button("generate_request"):
            doc = Document()
            doc.styles['Normal'].font.name = 'Times New Roman'
            doc.styles['Normal'].font.size = Pt(12)

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run1 = p.add_run('MIDWESTERN UMBRELLA OF WATER AND SANITATION\n')
            run1.bold = True
            run1.font.size = Pt(15)
            run1.font.color.rgb = RGBColor(0, 0, 0)

            run2 = p.add_run('MEMO\n')
            run2.bold = True
            run2.font.size = Pt(15)
            run2.font.color.rgb = RGBColor(0, 0, 0)

            run3 = p.add_run('='*54)
            run3.font.size = Pt(14)
            run3.bold = True
            run3.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Add left-aligned section with recipient info
            to_section = doc.add_paragraph()
            to_section.alignment = WD_ALIGN_PARAGRAPH.LEFT
            to_section.add_run("To:                     The Manager, MWUWS\n")
            to_section.add_run("Through:          The Senior Engineer, MWUWS\n")
            to_section.add_run("From:                The Mechanical Engineer, MWUWS\n")

            today_str = datetime.date.today().strftime("%d %B, %Y")
            doc.add_paragraph(f"Date: {today_str}")

            table = doc.add_table(rows=1, cols=len(edited_df.columns))
            table.style = 'Table Grid'

            hdr_cells = table.rows[0].cells
            for idx, col in enumerate(edited_df.columns):
                hdr_cells[idx].text = str(col)
                for paragraph in hdr_cells[idx].paragraphs:
                    paragraph.runs[0].font.bold = True
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            for _, row in edited_df.iterrows():
                row_cells = table.add_row().cells
                for idx, value in enumerate(row):
                    row_cells[idx].text = str(value)

            doc.add_paragraph("\nPrepared by: ___________________________")
            doc.add_paragraph("Approved by: ___________________________")

            doc_io = BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            st.download_button(
                label="Download Repair Request (Word)",
                data=doc_io,
                file_name=f"repair_request_{today_str}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

