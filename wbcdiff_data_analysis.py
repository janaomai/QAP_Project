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

# Removing outliers
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
    template_path = r"C:\iCCnet QAP Program\Source_files\WBCWordTemplate.docx"

    Database.columns = Database.columns.str.strip()

    bloodcells = ['wcc', 'neut', 'lymph', 'mono', 'eosino', 'baso']
    bloodcell_full_names = {
        'wcc': 'White Cell Count',
        'neut': 'Neutrophils',
        'lymph': 'Lymphocytes',
        'mono': 'Monocytes',
        'eosino': 'Eosinophils',
        'baso': 'Basophils'
    }

    bloodcell_units = {
        'wcc': 'x10⁹/L',
        'neut': '%',
        'lymph': '%',
        'mono': '%',
        'eosino': '%',
        'baso': '%'
    }

    for bloodcell in bloodcells:
        Database[bloodcell] = pd.to_numeric(Database[bloodcell], errors='coerce')

    Database_cleaned = remove_outliers(Database, bloodcells)

    for bloodcell in bloodcells[1:]:
        Database_cleaned[bloodcell + '_percent'] = (Database_cleaned[bloodcell] / Database_cleaned['wcc']) * 100
    
    medians = {bloodcell: np.median(Database_cleaned[bloodcell].dropna()) for bloodcell in bloodcells}
    medians['wcc'] = round(medians['wcc'], 1)  
    medians_percent = {bloodcell + '_percent': np.median(Database_cleaned[bloodcell + '_percent'].dropna()) for bloodcell in bloodcells[1:]}

    limits = {}
    for bloodcell, median in medians.items():
        if bloodcell == 'wcc':
            if median < 5.1:
                ll = round(median - 0.5)
                ul = round(median + 0.5)
            else: 
                ll = round(median * 0.9)
                ul = round(median * 1.1)
        limits[bloodcell] = (ll, ul)

    for bloodcell, median in medians_percent.items():
        if bloodcell == 'neut_percent':
            if median < 10.1:
                ll = round(median - 1, 1)
                ul = round(median + 1, 1)
            else: 
                ll = round(median * 0.9, 1)
                ul = round(median * 1.1, 1)
        elif bloodcell == 'lymph_percent':
            if median < 10.1:
                ll = round(median - 2, 1)
                ul = round(median + 2, 1)
            else:
                ll = round(median * 0.8, 1)
                ul = round(median * 1.2, 1)
        elif bloodcell == 'mono_percent':
            if median < 10.1:
                ll = round(median - 3, 1)
                ul = round(median + 3, 1)
            else:
                ll = round(median * 0.7, 1)
                ul = round(median * 1.3, 1)
        elif bloodcell == 'eosino_percent':
            if median < 10.1:
                ll = round(median - 3, 1)
                ul = round(median + 3, 1)
            else:
                ll = round(median * 0.7, 1)
                ul = round(median * 1.3, 1)
        elif bloodcell == 'baso_percent':
            if median < 10.1:
                ll = round(median - 3, 1)
                ul = round(median + 3, 1)
            else:
                ll = round(median * 0.7, 1)
                ul = round(median * 1.3, 1)
        limits[bloodcell] = (ll, ul)

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

        program_paragraph = doc.add_paragraph()
        run = program_paragraph.add_run("Program: White Blood Cell Differential")
        run.bold = True
        program_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        device_paragraph = doc.add_paragraph()
        run = device_paragraph.add_run("Device: WBC")
        run.bold = True
        device_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        site_paragraph = doc.add_paragraph()
        run = site_paragraph.add_run(site)
        run.bold = True
        site_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        sample_paragraph = doc.add_paragraph()
        run = sample_paragraph.add_run(f"Sample: {sheet_name}")
        run.bold = True
        sample_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Create a table with header row only
        table = doc.add_table(rows=1, cols=7)

        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Blood Cell'
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

        for i, bloodcell in enumerate(bloodcells):
            row_cells = table.add_row().cells
            bloodcell_run = row_cells[0].paragraphs[0].add_run(bloodcell_full_names[bloodcell])  
            bloodcell_run.font.bold = True
            your_result = row[bloodcell]
            if pd.isna(your_result):
                row_cells[1].text = 'No submission'
                row_cells[2].text = "{:2.f}"(limits[bloodcell][0])
                row_cells[3].text = "{:2.f}"(medians[bloodcell])
                row_cells[4].text = "{:2.f}"(limits[bloodcell][1])
                row_cells[5].text = bloodcell_units[bloodcell]  
            else:
                if bloodcell == 'wcc':
                    row_cells[1].text = "{:2.f}"(your_result)
                    row_cells[2].text = "{:2.f}"(limits[bloodcell][0])
                    row_cells[3].text = "{:2.f}"(medians[bloodcell])
                    row_cells[4].text = "{:2.f}"(limits[bloodcell][1])
                    row_cells[5].text = bloodcell_units[bloodcell]  
                else:
                    your_result_percent = round((your_result / row['wcc']) * 100, 1)
                    row_cells[1].text = str(round(your_result_percent, 1))
                    row_cells[2].text = str(round(limits[bloodcell + '_percent'][0], 1))
                    row_cells[3].text = str(round(medians_percent[bloodcell + '_percent'], 1))
                    row_cells[4].text = str(round(limits[bloodcell + '_percent'][1], 1))
                    row_cells[5].text = bloodcell_units[bloodcell]

            if bloodcell == 'wcc':
                if limits[bloodcell][0] <= your_result <= limits[bloodcell][1]:
                    row_cells[6].text = 'Acceptable'
                else:
                    row_cells[6].text = 'Unacceptable'
            else:
                if limits[bloodcell + '_percent'][0] <= your_result_percent <= limits[bloodcell + '_percent'][1]:
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

        fig, axes = plt.subplots(2, 3, figsize=(16, 10))
        fig.tight_layout(pad=10.0, h_pad=8)
        axes = axes.flatten()

        for idx, bloodcell in enumerate(bloodcells):  
            ax = axes[idx]

            if bloodcell == 'wcc':
                ax.plot(np.random.normal(1, 0.04, size=len(Database_cleaned[bloodcell].dropna())), Database_cleaned[bloodcell].dropna(), 'o', color='black', label="Other sites", markersize=8)
                ax.set_ylim(Database_cleaned[bloodcell].min() - 1, Database_cleaned[bloodcell].max() + 1)
                ax.set_ylabel(bloodcell_units[bloodcell], fontsize=18)
                ax.set_xlabel(bloodcell_full_names[bloodcell], fontsize=20, fontweight='bold', loc='left')
                your_result_wcc = row[bloodcell]
                ax.scatter(1, your_result_wcc, color='fuchsia', marker='s', label='Your result', s=80, zorder=2)
                ax.axhspan(limits[bloodcell][0], medians[bloodcell], color='green', alpha=0.2)
                ax.axhspan(medians[bloodcell], limits[bloodcell][1], color='green', alpha=0.2)

                if limits[bloodcell][0] <= your_result_wcc <= limits[bloodcell][1]:
                    ax.set_title('Acceptable', fontsize=17, fontweight='bold', color='green')
                else:
                    ax.set_title('Unacceptable', fontsize=17, fontweight='bold', color='red')

                ax.annotate('RCPA ALP: +/- 0.5 up to 5x10⁹/L\nthen 10%', xy=(0, -0.25), xycoords='axes fraction', fontsize=11, ha='left')

            else:
                ax.plot(np.random.normal(1, 0.04, size=len(Database_cleaned[bloodcell + '_percent'].dropna())), Database_cleaned[bloodcell + '_percent'].dropna(), 'o', color='black', label="Other sites", markersize=8)
                if bloodcell in ['mono', 'baso']:
                    ax.set_ylim(-1, 25)
                else: 
                    ax.set_ylim(-1, 100)
                your_result_percent = (row[bloodcell] / row['wcc']) * 100
                ax.scatter(1, your_result_percent, color='fuchsia', marker='s', label='Your result', s=80, zorder=2)
                ax.axhspan(limits[bloodcell + '_percent'][0], medians_percent[bloodcell + '_percent'], color='green', alpha=0.2)
                ax.axhspan(medians_percent[bloodcell + '_percent'], limits[bloodcell + '_percent'][1], color='green', alpha=0.2)

                if limits[bloodcell + '_percent'][0] <= your_result_percent <= limits[bloodcell + '_percent'][1]:
                    ax.set_title('Acceptable', fontsize=17, fontweight='bold', color='green')
                    ax.set_xlabel(bloodcell_full_names[bloodcell], fontsize=20, fontweight='bold', loc='left')  
                    ax.set_ylabel(bloodcell_units[bloodcell], fontsize=16) 
                else:
                    ax.set_title('Unacceptable', fontsize=17, fontweight='bold', color='red')
                    ax.set_xlabel(bloodcell_full_names[bloodcell], fontsize=20, fontweight='bold', loc='left')  
                    ax.set_ylabel(bloodcell_units[bloodcell], fontsize=16)  

                additional_text = {
                    'neut': 'RCPA ALP: +/- 1 up to 10%\nthen 10%',
                    'lymph': 'RCPA ALP: +/- 2 up to 10%\nthen 20%',
                    'mono': 'RCPA ALP: +/- 3 up to 10%\nthen 30%',
                    'eosino': 'RCPA ALP: +/- 3 up to 10%\nthen 30%',
                    'baso': 'RCPA ALP: +/- 3 up to 10%\nthen 30%'
                }
                ax.annotate(additional_text[bloodcell], xy=(0, -0.25), xycoords='axes fraction', fontsize=11, ha='left')

            ax.legend(loc='upper right', bbox_to_anchor=(1.45, 0), fontsize=13)
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
        output_path = os.path.join(output_dir, f"WBC_{site}_{today_date}.docx")
        doc.save(output_path)
