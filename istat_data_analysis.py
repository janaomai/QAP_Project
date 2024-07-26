# Libraries
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
    template_path = r"C:\iCCnet QAP Program\Source_files\iSTATWordTemplate.docx"

    Database.columns = Database.columns.str.strip()
    analytes = ['ph', 'pco2', 'po2', 'lac', 'na', 'k', 'ica', 'glu', 'urea', 'creat', 'hct']

    for analyte in analytes:
        Database[analyte] = pd.to_numeric(Database[analyte], errors='coerce')

    Database_cleaned = remove_outliers(Database, analytes)
    means = {analyte: np.mean(Database_cleaned[analyte].dropna()) for analyte in analytes}

    limits = {}
    for analyte, mean in means.items():
        if analyte == 'ph':
            ll = round(mean - 0.04, 2)
            ul = round(mean + 0.04, 2)
        elif analyte == 'pco2':
            if mean > 34:
                ll = round(mean * 0.94, 1)
                ul = round(mean * 1.06, 1)
            else:
                ll = round(mean - 2, 1)
                ul = round(mean + 2, 1)
        elif analyte == 'po2':
            if mean > 83:
                ll = round(mean * 0.94, 0)
                ul = round(mean * 1.06, 0)
            else:
                ll = round(mean - 5, 0)
                ul = round(mean + 5, 0)
        elif analyte == 'lac':
            if mean > 4.0:
                ll = round(mean * 0.88, 2)
                ul = round(mean * 1.12, 2)
            else:
                ll = round(mean - 0.5, 2)
                ul = round(mean + 0.5, 2)
        elif analyte == 'na':
            if mean > 150:
                ll = round(mean * 0.98, 0)
                ul = round(mean * 1.02, 0)
            else:
                ll = round(mean - 3, 0)
                ul = round(mean + 3, 0)
        elif analyte == 'k':
            if mean > 4:
                ll = round(mean * 0.95, 1)
                ul = round(mean * 1.05, 1)
            else:
                ll = round(mean - 0.2, 1)
                ul = round(mean + 0.2, 1)
        elif analyte == 'ica':
            if mean > 1:
                ll = round(mean * 0.96, 2)
                ul = round(mean * 1.04, 2)
            else:
                ll = round(mean - 0.04, 2)
                ul = round(mean + 0.04, 2)
        elif analyte == 'glu':
            if mean > 5:
                ll = round(mean * 0.92, 1)
                ul = round(mean * 1.08, 1)
            else:
                ll = round(mean - 0.4, 1)
                ul = round(mean + 0.4, 1)
        elif analyte == 'urea':
            if mean > 4.0:
                ll = round(mean * 0.88, 1)
                ul = round(mean * 1.12, 1)
            else:
                ll = round(mean - 0.5, 1)
                ul = round(mean + 0.5, 1)
        elif analyte == 'creat':
            if mean > 100:
                ll = round(mean * 0.92, 0)
                ul = round(mean * 1.08, 0)
            else:
                ll = round(mean - 8, 0)
                ul = round(mean + 8, 0)
        elif analyte == 'hct':
            if mean > 20:
                ll = round(mean * 0.80, 0)
                ul = round(mean * 1.2, 0)
            else:
                ll = round(mean - 4, 0)
                ul = round(mean + 4, 0)

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
        run = device_paragraph.add_run("Device: iSTAT")
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
        table = doc.add_table(rows=len(analytes) + 1, cols=7)

        # Header row names
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Analyte'
        hdr_cells[1].text = 'Your Result'
        hdr_cells[2].text = 'Lower Limit'
        hdr_cells[3].text = 'Mean'
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
            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
            shading = OxmlElement('w:shd')
            shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', '800000')  # Dark red
            cell._element.get_or_add_tcPr().append(shading)

        # Data row values
        analyte_names = {
            'ph': 'pH',
            'pco2': 'pCO2',
            'po2': 'pO2',
            'lac': 'Lactate',
            'na': 'Sodium',
            'k': 'Potassium',
            'ica': 'Ionised Calcium',
            'glu': 'Glucose',
            'urea': 'Urea',
            'creat': 'Creatinine',
            'hct': 'Haematocrit'
        }

        y_labels = {
            'ph': '',
            'pco2': 'mmHg',
            'po2': 'mmHg',
            'lac': 'mmol/L',
            'na': 'mmol/L',
            'k': 'mmol/L',
            'ica': 'mmol/L',
            'glu': 'mmol/L',
            'urea': 'mmol/L',
            'creat': 'µmol/L',
            'hct': 'Percent (%)'
        }

        for i, analyte in enumerate(analytes):
            row_cells = table.rows[i + 1].cells
            analyte_run = row_cells[0].paragraphs[0].add_run(analyte_names[analyte])
            analyte_run.font.bold = True
            your_result = site_data[analyte].values[0]
            if pd.isna(your_result):
                row_cells[1].text = 'No submission'
                row_cells[2].text = str(round(limits[analyte][0], 2))
                row_cells[3].text = str(round(means[analyte], 2))
                row_cells[4].text = str(round(limits[analyte][1], 2))
                row_cells[5].text = y_labels[analyte]
            else:
                row_cells[1].text = str(round(your_result, 2))
                row_cells[2].text = str(round(limits[analyte][0], 2))
                row_cells[3].text = str(round(means[analyte], 2))
                row_cells[4].text = str(round(limits[analyte][1], 2))
                row_cells[5].text = y_labels[analyte]
            # Interpretation based on limits
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

                # Alternate row colors
                shading = OxmlElement('w:shd')
                if i % 2 == 0:
                    shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', 'FCF7EC')  # Beige
                else:
                    shading.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill', 'FFFFFF')  # White
                cell._element.get_or_add_tcPr().append(shading)

        # Page break after table
        doc.add_page_break()

        # Graphing
        # Adjust the number of rows and columns dynamically
        num_analytes = len(analytes)
        rows = 4  # Fixed 4 rows
        cols = 3  # Fixed 3 columns

        # Graphing
        fig, axes = plt.subplots(rows, cols, figsize=(22, 22))
        fig.tight_layout(pad=9.0, h_pad=11)
        axes = axes.flatten()

        for idx, analyte in enumerate(analytes):
            ax = axes[idx]

            # Jitter plot
            ax.plot(np.random.normal(1, 0.04, size=len(Database[analyte].dropna())), Database[analyte].dropna(), 'o', color='black', label="Other sites", markersize=10)

            # y-limits
            if analyte == 'ph':
                ax.set_ylim(Database[analyte].min() - 0.2, Database[analyte].max() + 0.2)
            elif analyte == 'pco2':
                ax.set_ylim(Database[analyte].min() - 5, Database[analyte].max() + 5)
            elif analyte == 'po2':
                ax.set_ylim(Database[analyte].min() - 20, Database[analyte].max() + 20)
            elif analyte =='na':
                ax.set_ylim(Database[analyte].min() - 5, Database[analyte].max() + 5)
            elif analyte == 'k':
                ax.set_ylim(Database[analyte].min() - 1, Database[analyte].max() + 1)
            elif analyte == 'ica':
                ax.set_ylim(Database[analyte].min() - 0.2, Database[analyte].max() + 0.5)
            elif analyte == 'hct':
                ax.set_ylim(Database[analyte].min() - 10, Database[analyte].max() + 10)
            elif analyte =='glu':
                ax.set_ylim(Database[analyte].min() -1, Database[analyte].max() + 1)
            elif analyte == 'lac':
                ax.set_ylim(Database[analyte].min() - 0.5, Database[analyte].max() + 0.5)
            elif analyte == 'urea':
                ax.set_ylim(Database[analyte].min() - 1, Database[analyte].max() + 1)
            elif analyte == 'crea':
                ax.set_ylim(Database[analyte].min() - 10, Database[analyte].max() + 10)
            else:
                ax.set_ylim(Database[analyte].min() - 5, Database[analyte].max() + 5)

            # Dot for your result
            your_result = site_data[analyte].values[0]
            if pd.isna(your_result):
                ax.scatter([], [], color='fuchsia', marker='s', label='Your result', s=100, zorder=2)  # Do not plot
            else:
                ax.scatter(1, your_result, color='fuchsia', marker='s', label='Your result', s=80, zorder=2)

            # Shaded background for ll
            ax.axhspan(limits[analyte][0], means[analyte], color='green', alpha=0.2, zorder=0)

            # Shaded background for ul
            ax.axhspan(means[analyte], limits[analyte][1], color='green', alpha=0.2, zorder=0)

            # Print unacceptable results
            if not pd.isna(your_result) and not (limits[analyte][0] <= your_result <= limits[analyte][1]):
                print(f"{analyte} for site {site} is an unacceptable result: {your_result}")

            # Print result as acceptable or unacceptable for graph title
            if pd.isna(your_result):
                ax.set_title('Not Provided', fontsize=17, fontweight='bold', color='gray')
            elif limits[analyte][0] <= your_result <= limits[analyte][1]:
                ax.set_title('Acceptable', fontsize=22, fontweight='bold', color='green')
            else:
                ax.set_title('Unacceptable', fontsize=22, fontweight='bold', color='red')

            # Title and labels
            ax.set_xlabel(analyte_names[analyte], fontsize=28, fontweight='bold', loc='left')
            ax.set_ylabel(y_labels[analyte], fontsize=20)

            # Legend
            ax.legend(loc='upper right', bbox_to_anchor=(1.3, 0), fontsize=18)

            # Add additional lines of text below the x-axis label
            additional_text = {
                'ph': 'RCPA ALP: +/- 0.04',
                'pco2': 'RCPA ALP: +/- 2.0 up to 34 mmHg then 6%',
                'po2': 'RCPA ALP: +/- 5 up to 83 mmHg then 6%',
                'lac': 'RCPA ALP: +/- 0.5 up to 4 mmol/L then 12%',
                'na': 'RCPA ALP: +/- 3 up to 150 mmol/L then 2%',
                'k': 'RCPA ALP: +/- 0.2 up to 4 mmol/L then 5%',
                'ica': 'RCPA ALP: +/- 0.04 up to 1 mmol/L then 4%',
                'glu': 'RCPA ALP: +/- 0.4 up to 5 mmol/L then 8%',
                'urea': 'RCPA ALP: +/- 0.5 up to 4 mmol/L then 12%',
                'creat': 'RCPA ALP: +/- 8 up to 100 µmol/L then 8%',
                'hct': 'RCPA ALP: +/- 4 up to 20% then 20%'
            }
            # Add linespacing and place to left
            ax.annotate(additional_text[analyte], xy=(0, -0.18), xycoords='axes fraction', fontsize=14, ha='left')

            ax.set_xticks([])
            ax.tick_params(axis='y', labelsize=18)

        # Remove unused axes
        if len(analytes) < len(axes):
            for j in range(len(analytes), len(axes)):
                fig.delaxes(axes[j])

        # Save the figure to a temporary file
        plot_filename = f'{site}_combined.png'
        plt.savefig(plot_filename, bbox_inches='tight')
        plt.close()

        # Add the figure to the Word document using a table to control alignment
        graph_table = doc.add_table(rows=1, cols=1)
        graph_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        cell = graph_table.cell(0, 0)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(plot_filename, width=Inches(7))

        # Remove the temporary plot file
        os.remove(plot_filename)

        # Save the Word document for the site
        output_path = f'C:\\iCCnet QAP Program\\Output\POCT\\iSTAT_{site}_{today_date}.docx'
        doc.save(output_path)
