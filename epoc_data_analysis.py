import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os
from datetime import datetime
import main

# Globals to be set by main.py
file_path = ""
sheet_name = ""
user_id = ""

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
    template_path = r"C:\iCCnet QAP Program\Source_files\EpocWordTemplate.docx"

    Database.columns = Database.columns.str.strip()
    analytes = ['ph', 'pco2', 'po2', 'na', 'k', 'ica', 'cl', 'hct', 'glu', 'lac', 'urea', 'creat']

    for analyte in analytes:
        Database[analyte] = pd.to_numeric(Database[analyte], errors='coerce')

    Database_cleaned = remove_outliers(Database, analytes)
    medians = {analyte: np.median(Database_cleaned[analyte].dropna()) for analyte in analytes}

    limits = {}
    for analyte, median in medians.items():
        if analyte == 'ph':
            ll = round(median - 0.04, 2)
            ul = round(median + 0.04, 2)
        elif analyte == 'pco2':
            if median > 34:
                ll = round(median * 0.94, 1)
                ul = round(median * 1.06, 1)
            else:
                ll = round(median - 2, 1)
                ul = round(median + 2, 1)
        elif analyte == 'po2':
            if median > 83:
                ll = round(median * 0.94, 0)
                ul = round(median * 1.06, 0)
            else:
                ll = round(median - 5, 0)
                ul = round(median + 5, 0)
        elif analyte == 'na':
            if median > 150:
                ll = round(median * 0.98, 0)
                ul = round(median * 1.02, 0)
            else:
                ll = round(median - 3, 0)
                ul = round(median + 3, 0)
        elif analyte == 'k':
            if median > 4:
                ll = round(median * 0.95, 1)
                ul = round(median * 1.05, 1)
            else:
                ll = round(median - 0.2, 1)
                ul = round(median + 0.2, 1)
        elif analyte == 'ica':
            if median > 1:
                ll = round(median * 0.96, 2)
                ul = round(median * 1.04, 2)
            else:
                ll = round(median - 0.04, 2)
                ul = round(median + 0.04, 2)
        elif analyte == 'cl':
            if median > 100:
                ll = round(median * 0.97, 0)
                ul = round(median * 1.03, 0)
            else:
                ll = round(median - 3, 0)
                ul = round(median + 3, 0)
        elif analyte == 'hct':
            if median > 20:
                ll = round(median * 0.80, 0)
                ul = round(median * 1.2, 0)
            else:
                ll = round(median - 4, 0)
                ul = round(median + 4, 0)
        elif analyte == 'glu':
            if median > 5:
                ll = round(median * 0.92, 1)
                ul = round(median * 1.08, 1)
            else:
                ll = round(median - 0.4, 1)
                ul = round(median + 0.4, 1)
        elif analyte == 'lac':
            if median > 4.0:
                ll = round(median * 0.88, 2)
                ul = round(median * 1.12, 2)
            else:
                ll = round(median - 0.5, 2)
                ul = round(median + 0.5, 2)
        elif analyte == 'urea':
            if median > 4.0:
                ll = round(median * 0.88, 1)
                ul = round(median * 1.12, 1)
            else:
                ll = round(median - 0.5, 1)
                ul = round(median + 0.5, 1)
        elif analyte == 'creat':
            if median > 100:
                ll = round(median * 0.92, 0)
                ul = round(median * 1.08, 0)
            else:
                ll = round(median - 8, 0)
                ul = round(median + 8, 0)

        limits[analyte] = (ll, ul)

    for site in Database['site'].unique():
        site_data = Database[Database['site'] == site]
        if site_data.empty:
            continue

        doc = Document(template_path)
        today_date = datetime.now().strftime('%d-%m-%Y')
        replace_text(doc, 'DATE', today_date)
        replace_text(doc, 'SITE', site)
        replace_text(doc, 'CYCLE', sheet_name)
        replace_placeholder_in_footer(doc, 'ISSUER', user_id.title())

        # Add the program information as a bold and centered paragraph
        program_paragraph = doc.add_paragraph()
        run = program_paragraph.add_run("Program: Blood Gas & Electrolytes")
        run.bold = True
        program_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add the device information as a bold and centered paragraph
        device_paragraph = doc.add_paragraph()
        run = device_paragraph.add_run("Device: Epoc")
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

        # Add the metrics table
        table = doc.add_table(rows=1, cols=7)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Analyte'
        hdr_cells[1].text = 'Your Result'
        hdr_cells[2].text = 'Lower Limit'
        hdr_cells[3].text = 'Median'
        hdr_cells[4].text = 'Upper Limit'
        hdr_cells[5].text = 'Units'
        hdr_cells[6].text = 'Interpretation'

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

        analyte_names = {
            'ph': 'pH',
            'pco2': 'pCO2',
            'po2': 'pO2',
            'na': 'Sodium',
            'k': 'Potassium',
            'ica': 'Ionised Calcium',
            'cl': 'Chloride',
            'hct': 'Haematocrit',
            'glu': 'Glucose',
            'lac': 'Lactate',
            'urea': 'Urea',
            'creat': 'Creatinine'
        }
        y_labels = {
            'ph': '',
            'pco2': 'mmHg',
            'po2': 'mmHg',
            'na': 'mmol/L',
            'k': 'mmol/L',
            'ica': 'mmol/L',
            'cl': 'mmol/L',
            'hct': '%',
            'glu': 'mmol/L',
            'lac': 'mmol/L',
            'urea': 'mmol/L',
            'creat': 'µmol/L'
        }

        for i, analyte in enumerate(analytes):
            row_cells = table.add_row().cells
            analyte_run = row_cells[0].paragraphs[0].add_run(analyte_names[analyte])
            analyte_run.font.bold = True
            your_result = site_data[analyte].values[0]
            if pd.isna(your_result):
                row_cells[1].text = 'No submission'
                row_cells[2].text = str(round(limits[analyte][0], 2))
                row_cells[3].text = str(round(medians[analyte], 2))
                row_cells[4].text = str(round(limits[analyte][1], 2))
                row_cells[5].text = y_labels[analyte]
            else:
                row_cells[1].text = str(round(your_result, 2))
                row_cells[2].text = str(round(limits[analyte][0], 2))
                row_cells[3].text = str(round(medians[analyte], 2))
                row_cells[4].text = str(round(limits[analyte][1], 2))
                row_cells[5].text = y_labels[analyte]
            if pd.isna(your_result):
                row_cells[6].text = 'Unacceptable'
            elif limits[analyte][0] <= your_result <= limits[analyte][1]:
                run = row_cells[6].paragraphs[0].add_run('Acceptable')
                run.font.color.rgb = RGBColor(0, 128, 0)
            else:
                run = row_cells[6].paragraphs[0].add_run('Unacceptable')
                run.font.color.rgb = RGBColor(255, 0, 0)

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

        doc.add_page_break()

        fig, axes = plt.subplots(4, 3, figsize=(22, 22))
        fig.tight_layout(pad=9.0, h_pad=14)
        axes = axes.flatten()

        for idx, analyte in enumerate(analytes):
            ax = axes[idx]
 
            ax.plot(np.random.normal(1, 0.04, size=len(Database_cleaned[analyte].dropna())), Database_cleaned[analyte].dropna(), 'o', color='black', label="Other sites", markersize=8)
 
            if analyte == 'ph':
                ax.set_ylim(Database_cleaned[analyte].min() - 0.2, Database_cleaned[analyte].max() + 0.2)
            elif analyte == 'pco2':
                ax.set_ylim(Database_cleaned[analyte].min() - 5, Database_cleaned[analyte].max() + 5)
            elif analyte == 'po2':
                ax.set_ylim(Database_cleaned[analyte].min() - 20, Database_cleaned[analyte].max() + 20)
            elif analyte =='na':
                ax.set_ylim(Database_cleaned[analyte].min() - 5, Database_cleaned[analyte].max() + 5)
            elif analyte == 'k':
                ax.set_ylim(Database_cleaned[analyte].min() - 1, Database_cleaned[analyte].max() + 1)
            elif analyte == 'ica':
                ax.set_ylim(Database_cleaned[analyte].min() - 0.2, Database_cleaned[analyte].max() + 0.5)
            elif analyte =='cl':
                ax.set_ylim(Database_cleaned[analyte].min() - 5, Database_cleaned[analyte].max() + 5)
            elif analyte == 'hct':
                ax.set_ylim(Database_cleaned[analyte].min() - 10, Database_cleaned[analyte].max() + 10)
            elif analyte =='glu':
                ax.set_ylim(Database_cleaned[analyte].min() -1, Database_cleaned[analyte].max() + 1)
            elif analyte == 'lac':
                ax.set_ylim(Database_cleaned[analyte].min() - 0.5, Database_cleaned[analyte].max() + 0.5)
            elif analyte == 'urea':
                ax.set_ylim(Database_cleaned[analyte].min() - 1, Database_cleaned[analyte].max() + 1)
            elif analyte == 'crea':
                ax.set_ylim(Database_cleaned[analyte].min() - 10, Database_cleaned[analyte].max() + 10)
            else:
                ax.set_ylim(Database_cleaned[analyte].min() - 5, Database_cleaned[analyte].max() + 5)

            your_result = site_data[analyte].values[0]
            ax.scatter(1, your_result, color='fuchsia', marker='s', label='Your result', s=80, zorder=2)

            ax.axhspan(limits[analyte][0], medians[analyte], color='green', alpha=0.2, zorder=0)
            ax.axhspan(medians[analyte], limits[analyte][1], color='green', alpha=0.2, zorder=0)

            if not (limits[analyte][0] <= your_result <= limits[analyte][1]):
                print(f"{analyte} for site {site} is an unacceptable result: {your_result}")

            if limits[analyte][0] <= your_result <= limits[analyte][1]:
                ax.set_title('Acceptable', fontsize=22, fontweight='bold', color='green')
            else:
                ax.set_title('Unacceptable', fontsize=22, fontweight='bold', color='red')

            ax.set_xlabel(analyte_names[analyte], fontsize=28, fontweight='bold', loc='left')
            ax.set_ylabel(y_labels[analyte], fontsize=20)

            ax.legend(loc='upper right', bbox_to_anchor=(1.2, 0), fontsize=17)

            additional_text = {
                'ph': 'RCPA ALP: +/- 0.04',
                'pco2': 'RCPA ALP: +/- 2.0 up to 34 mmHg then 6%',
                'po2': 'RCPA ALP: +/- 5 up to 83 mmHg then 6%',
                'na': 'RCPA ALP: +/- 3 up to 150 mmol/L then 2%',
                'k': 'RCPA ALP: +/- 0.2 up to 4 mmol/L then 5%',
                'ica': 'RCPA ALP: +/- 0.04 up to 1 mmol/L then 4%',
                'cl': 'RCPA ALP: +/- 3 up to 100 mmol/L then 3%',
                'hct': 'RCPA ALP: +/- 4 up to 20% then 20%',
                'glu': 'RCPA ALP: +/- 0.4 up to 5 mmol/L then 8%',
                'lac': 'RCPA ALP: +/- 0.5 up to 4 mmol/L then 12%',
                'urea': 'RCPA ALP: +/- 0.5 up to 4 mmol/L then 12%',
                'creat': 'RCPA ALP: +/- 8 up to 100 µmol/L then 8%'
            }
            ax.annotate(additional_text[analyte], xy=(0, -0.2), xycoords='axes fraction', fontsize=13, ha='left')

            ax.set_xticks([])

            ax.tick_params(axis='y', labelsize=20)

        output_dir = r"C:\iCCnet QAP Program\Output\POCT"
        plot_filename = os.path.join(output_dir, f'{site}_combined.png')
        plt.savefig(plot_filename, bbox_inches='tight')
        plt.close()

        graph_table = doc.add_table(rows=1, cols=1)
        graph_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        cell = graph_table.cell(0, 0)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(plot_filename, width=Inches(7.2))

        os.remove(plot_filename)

        today_date = datetime.now().strftime('%d-%m-%Y')
        output_path = os.path.join(output_dir, f"Epoc_{site}_{today_date}.docx")
        doc.save(output_path)
