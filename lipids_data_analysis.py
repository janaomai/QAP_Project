# Libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
from datetime import datetime

# Globals to be set by main.py
file_path = ""
sheet_name = ""
user_id = ""

# Replacing text in template doc
def replace_text(doc, search_text, replace_text):
    for paragraph in doc.paragraphs:
        if search_text in paragraph.text:
            paragraph.text = paragraph.text.replace(search_text, replace_text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if search_text in cell.text:
                    cell.text = cell.text.replace(search_text, replace_text)

def replace_placeholder_in_footer(doc, placeholder, replacement_text):
    for section in doc.sections:
        for paragraph in section.footer.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement_text)
                for run in paragraph.runs:
                    if replacement_text in run.text:
                        run.font.size = Pt(7)
                        run.font.bold = False

def remove_outliers(data, analytes):
    cleaned_data = data.copy()
    for analyte in analytes:
        Q1 = cleaned_data[analyte].quantile(0.15)
        Q3 = cleaned_data[analyte].quantile(0.85)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR
        cleaned_data = cleaned_data[(cleaned_data[analyte] >= lower_bound) & (cleaned_data[analyte] <= upper_bound)]
    return cleaned_data

def run():
    global file_path, sheet_name, user_id

    if not file_path:
        raise FileNotFoundError("No file path specified.")
    if not sheet_name:
        raise ValueError("No sheet name specified.")
    if not user_id:
        raise ValueError("No user ID specified.")

    Database = pd.read_excel(file_path, sheet_name=sheet_name)
    template_path = r"C:\iCCnet QAP Program\Source_files\LipidsWordTemplate.docx"

    Database.columns = Database.columns.str.strip()
    analytes = ['chol', 'ldl', 'hdl', 'trig']

    # Re-naming analytes for table
    analyte_names = {'chol': 'Cholesterol',
                     'ldl': 'LDL',
                     'hdl': 'HDL',
                     'trig': 'Triglycerides'}

    for analyte in analytes:
        Database[analyte] = pd.to_numeric(Database[analyte], errors='coerce')

    Database_cleaned = remove_outliers(Database, analytes)
    medians = {analyte: np.median(Database_cleaned[analyte].dropna()) for analyte in analytes}

    limits = {}
    for analyte, median in medians.items():
        if analyte == 'chol':
            if median > 5:
                ll = round(median * 0.94, 2)
                ul = round(median * 1.06, 2)
            else: 
                ll = round(median - 0.30, 2)
                ul = round(median + 0.30, 2)
        elif analyte == 'ldl':
            if median > 2:
                ll = round(median * 0.90, 2)
                ul = round(median * 1.10, 2)
            else: 
                ll = round(median - 0.20, 2)
                ul = round(median + 0.20, 2)
        elif analyte == 'hdl':
            if median > 0.8:
                ll = round(median * 0.88, 2)
                ul = round(median * 1.12, 2)
            else:
                ll = round(median - 0.10, 2)
                ul = round(median + 0.10, 2)
        elif analyte == 'trig':
            if median > 1.60:
                ll = round(median * 0.88, 2)
                ul = round(median * 1.12, 2)
            else: 
                ll = round(median - 0.20, 2)
                ul = round(median + 0.20, 2)

        limits[analyte] = (ll, ul)

    for index, row in Database.iterrows():
        site = row['site']
        print(f"Processing site: {site}")
        if pd.isnull(site):
            continue

        doc = Document(template_path)

        today_date = datetime.now().strftime('%d/%m/%Y')
        replace_text(doc, 'DATE', today_date)
        replace_text(doc, 'SITE', site)
        replace_text(doc, 'CYCLE', sheet_name)
        replace_placeholder_in_footer(doc, 'ISSUER', user_id.title())

        # Add the program information as a bold and centered paragraph
        program_paragraph = doc.add_paragraph()
        run = program_paragraph.add_run("Program: Lipids")
        run.bold = True
        program_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add the device information as a bold and centered paragraph
        device_paragraph = doc.add_paragraph()
        run = device_paragraph.add_run("Device: Cobas b101")
        run.bold = True
        device_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add the site name as a bold and centered paragraph
        site_paragraph = doc.add_paragraph()
        run = site_paragraph.add_run(site)
        run.bold = True
        site_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add the sample name as a bold and centered paragraph
        sample_paragraph = doc.add_paragraph()
        run = sample_paragraph.add_run(f"Sample: {sheet_name}")
        run.bold = True
        sample_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Table
        table = doc.add_table(rows=1, cols=7)

        # Header row names
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Analyte'
        hdr_cells[1].text = 'Your Result'
        hdr_cells[2].text = 'Lower Limit'
        hdr_cells[3].text = 'Median'
        hdr_cells[4].text = 'Upper Limit'
        hdr_cells[5].text = 'Units'
        hdr_cells[6].text = 'Interpretation'

        # Header styling
        for cell in hdr_cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if not cell.paragraphs[0].runs:
                cell.paragraphs[0].add_run()
            run = cell.paragraphs[0].runs[0]
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            shading = OxmlElement('w:shd')
            shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', '800000')
            cell._element.get_or_add_tcPr().append(shading)

        for i, analyte in enumerate(analytes):
            row_cells = table.add_row().cells
            analyte_run = row_cells[0].paragraphs[0].add_run(analyte_names[analyte])
            analyte_run.font.bold = True
            your_result = row[analyte]
            if pd.isna(your_result):
                row_cells[1].text = 'No submission'
                row_cells[2].text = "{:.2f}".format(limits[analyte][0])
                row_cells[3].text = "{:.2f}".format(medians[analyte])
                row_cells[4].text = "{:.2f}".format(limits[analyte][1])
                row_cells[5].text = 'mmol/L'
            else:
                row_cells[1].text = "{:.2f}".format(your_result)
                row_cells[2].text = "{:.2f}".format(limits[analyte][0])
                row_cells[3].text = "{:.2f}".format(medians[analyte])
                row_cells[4].text = "{:.2f}".format(limits[analyte][1])
                row_cells[5].text = 'mmol/L'
            if limits[analyte][0] <= your_result <= limits[analyte][1]:
                row_cells[6].text = 'Acceptable'
            else:
                row_cells[6].text = 'Unacceptable'

            for cell in row_cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if not cell.paragraphs[0].runs:
                    cell.paragraphs[0].add_run()
                run = cell.paragraphs[0].runs[0]
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 0, 0)
                shading = OxmlElement('w:shd')
                if i % 2 == 0:
                    shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', 'FCF7EC')
                else:
                    shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', 'FFFFFF')
                cell._element.get_or_add_tcPr().append(shading)

        doc.add_paragraph()

        fig, axes = plt.subplots(2, 2, figsize=(13.5, 9))
        fig.tight_layout(pad=6.0, h_pad=8)  
        axes = axes.flatten()

        for idx, analyte in enumerate(analytes):
            ax = axes[idx]
            ax.plot(np.random.normal(1, 0.04, size=len(Database[analyte].dropna())), Database[analyte].dropna(), 'o', color='black', label="Other sites", markersize=8)
            ax.set_ylim(Database[analyte].min() - 0.5, Database[analyte].max() + 0.5)
            your_result = row[analyte]
            ax.scatter(1, your_result, color='fuchsia', marker='s', label='Your result', s=80, zorder=2)
            ax.axhspan(limits[analyte][0], medians[analyte], color='green', alpha=0.2)
            ax.axhspan(medians[analyte], limits[analyte][1], color='green', alpha=0.2)
            if not (limits[analyte][0] <= your_result <= limits[analyte][1]):
                print(f"{analyte} for site {site} is an unacceptable result: {your_result}")
            if limits[analyte][0] <= your_result <= limits[analyte][1]:
                ax.set_title('Acceptable', fontsize=17, fontweight='bold', color='green')
            else:
                ax.set_title('Unacceptable', fontsize=17, fontweight='bold', color='red')
            ax.set_xlabel(analyte_names[analyte], fontsize=20, fontweight='bold', loc='left')
            ax.set_ylabel('mmol/L', fontsize=16)
            ax.legend(loc='upper right', bbox_to_anchor=(1.1, 0), fontsize=14)
            additional_text = {
                'chol': 'RCPA ALP: +/- 0.3 up to 5 mmol/L then 6%',
                'ldl': 'RCPA ALP: +/- 0.2 up to 2 mmol/L then 10%',
                'hdl': 'RCPA ALP: +/- 0.1 up to 0.8 mmol/L then 12%',
                'trig': 'RCPA ALP: +/- 0.2 up to 1.6 mmol/L then 12%'
            }
            ax.annotate(additional_text[analyte], xy=(0, -0.18), xycoords='axes fraction', fontsize=11, ha='left')
            ax.set_xticks([])
            ax.tick_params(axis='y', labelsize=15)

        output_dir = r"C:\iCCnet QAP Program\Output\POCT"
        plot_filename = os.path.join(output_dir, f'{site}_combined.png')
        plt.savefig(plot_filename, bbox_inches='tight')
        plt.close()

        graph_table = doc.add_table(rows=1, cols=1)
        graph_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        cell = graph_table.cell(0, 0)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(plot_filename, width=Inches(7))

        os.remove(plot_filename)

        today_date = datetime.now().strftime('%d-%m-%Y')
        output_path = os.path.join(output_dir, f"Lipids_{site}_{today_date}.docx")
        doc.save(output_path)