import matplotlib
matplotlib.use('Agg')

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO

def calculate_limits(analyte, data):
    numeric_data = pd.to_numeric(data, errors='coerce').dropna()
    median = numeric_data.median()
    if analyte == "D-Dimer":
        lower_limit = median * 0.50 if median > 0.50 else median - 0.25
        upper_limit = median * 1.50 if median > 0.50 else median + 0.25
    elif analyte == "CRP":
        lower_limit = median - 0.8 if median < 4.0 else median * 0.80
        upper_limit = median + 0.8 if median < 4.0 else median * 1.20
    elif analyte == "INR":
        lower_limit = median - 0.3 if median < 2.0 else median * 0.85
        upper_limit = median + 0.3 if median < 2.0 else median * 1.15
    elif analyte == "Troponin":
        lower_limit = median - 10 if median < 50 else median * 0.80
        upper_limit = median + 10 if median < 50 else median * 1.20
    elif analyte == "NT-ProBNP":
        lower_limit = median - 25 if median < 125 else median * 0.80
        upper_limit = median + 25 if median < 125 else median * 1.20
    elif analyte == "Haemoglobin":
        lower_limit = median - 5 if median < 100 else median * 0.95
        upper_limit = median + 5 if median < 100 else median * 1.05
    elif analyte == "HbA1c":
        lower_limit = median - 4 if median < 45 else median * 0.92
        upper_limit = median + 4 if median < 45 else median * 1.08
    elif analyte == "Glucose":
        lower_limit = median - 0.5 if median < 5 else median * 0.90
        upper_limit = median + 0.5 if median < 5 else median * 1.10
    elif analyte == "Ketones":
        lower_limit = median - 0.5 if median < 5 else median * 0.90
        upper_limit = median + 0.5 if median < 5 else median * 1.10
    else:
        raise ValueError(f"Unknown analyte: {analyte}")
    return lower_limit, upper_limit, median

def evaluate_results(sites, results, cycle, analyte):
    lower_limit, upper_limit, median = calculate_limits(analyte, results)
    interpretations = []
    for site, result in zip(sites, results):
        if result == "No submission" or result == "Unacceptable":
            interpretation = "Unacceptable"
        else:
            result = float(result)  # Ensure the result is numeric
            if lower_limit <= result <= upper_limit:
                interpretation = "Acceptable"
            else:
                interpretation = "Unacceptable"
        interpretations.append((site, lower_limit, upper_limit, median, result, interpretation))
    return interpretations

def process_metrics(sites, results, cycle, analyte):
    try:
        results = evaluate_results(sites, results, cycle, analyte)
        if not results:
            raise ValueError("No results returned from evaluate_results")
        histograms = {}
        format_mapping = {
            "D-Dimer": "{:.2f}",
            "CRP": "{:.1f}",
            "INR": "{:.1f}",
            "Troponin": "{:.0f}",
            "NT-ProBNP": "{:.0f}",
            "Haemoglobin": "{:.0f}",
            "HbA1c": "{:.1f}",
            "Glucose": "{:.1f}",
            "Ketones": "{:.1f}"
        }
        
        format_string = format_mapping.get(analyte, "{:.2f}")
        
        for site, lower_limit, upper_limit, median, result, interpretation in results:
            lower_limit = float(format_string.format(lower_limit))
            upper_limit = float(format_string.format(upper_limit))
            median = float(format_string.format(median))
            
            img_stream = BytesIO()
            plot_histogram(site, results, analyte, lower_limit, upper_limit, median, result, interpretation, cycle, img_stream, format_string)
            histograms[site] = img_stream
        
        return results, histograms  
    except Exception as e:
        print(f"Error processing metrics: {str(e)}")
        return None, None

def get_analyte_label(analyte):
    labels = {
        "D-Dimer": "ug/mL",
        "CRP": "mg/L",
        "INR": "INR",
        "Troponin": "ng/L",
        "NT-ProBNP": "pg/mL",
        "Haemoglobin": "g/L",
        "HbA1c": "mmol/mol",
        "Glucose": "mmol/L",
        "Ketones": "mmol/L"
    }
    return labels.get(analyte, "Concentration")

def get_analyte_sublabel(analyte):
    sublabels = {
        "D-Dimer": "RCPA ALP: Median +/- 0.25 up to 0.5 then +/- 50%",
        "CRP": "RCPA ALP: Median +/- 0.8 up to 4.0 then +/- 20%",
        "INR": "RCPA ALP: Median +/- 0.3 up to 2.0 then +/- 15%",
        "Troponin": "RCPA: ALP Median+/- 10 up to 50 then +/- 20%",
        "NT-ProBNP": "RCPA ALP: Median +/- 25 up to 125pg/mL then +/- 20%",
        "Haemoglobin": "RCPA ALP: Median +/- 5 up to 100g/L then +/- 5%",
        "HbA1c": "RCPA ALP: Median +/- 4 up to 45 mmol/mol then +/- 8%",
        "Glucose": "RCPA ALP: Median +/- 0.5 up to 5.0 then +/- 10%",
        "Ketones": "RCPA ALP: Median +/- 0.5 up to 5.0 then +/- 10%"
    }
    return sublabels.get(analyte, "")

def plot_histogram(site, results, analyte, lower_limit, upper_limit, median, site_result, interpretation, cycle, img_stream, format_string):
    numeric_results = [res[4] for res in results if isinstance(res[4], (int, float))]

    fig, ax = plt.subplots(figsize=(7, 5))

    if analyte == "INR":
        min_bin = np.floor(lower_limit * 10) / 10
        max_bin = np.ceil(upper_limit * 10) / 10
        bin_width = 0.1
        bins = np.arange(min_bin, max_bin + bin_width, bin_width)
        rect_size = upper_limit * 0.12  # 12% extension for INR
        lower_limit_rect_start = lower_limit - rect_size
        upper_limit_rect_end = upper_limit + rect_size
        x_ticks = np.arange(lower_limit_rect_start, upper_limit_rect_end + bin_width, bin_width)
    else:
        max_result = max(max(numeric_results), upper_limit) * 1.05
        step = upper_limit * 0.026
        lower_limit_extension = lower_limit * 0.10  # 20% extension for non-INR analytes
        bins = np.arange(lower_limit - lower_limit_extension, max_result + step, step)
        lower_limit_rect_start = lower_limit - lower_limit_extension
        upper_limit_rect_end = upper_limit + upper_limit * 0.10
        x_ticks = bins

    n, bins, patches = ax.hist(numeric_results, bins=bins, edgecolor='black')

    for patch in patches:
        patch.set_facecolor('red')

    for patch, bin_left in zip(patches, bins[:-1]):
        if lower_limit <= bin_left + (bins[1] - bins[0]) / 2 <= upper_limit:
            patch.set_facecolor('lightgreen')

    if isinstance(site_result, (int, float)):
        ax.axvline(x=site_result, color='purple', linewidth=3)  # Increase the linewidth to make the line bolder
        # Annotate 'Your result' at a fixed vertical position
        ax.annotate('Your result', xy=(site_result, 0.98), xycoords=('data', 'axes fraction'), 
                    xytext=(0, 10), textcoords='offset points', ha='center', color='black', fontsize=12, weight='bold')
        ax.plot(site_result, 1.00, marker='v', markersize=30, color='purple', transform=ax.get_xaxis_transform())

    if analyte == "INR":
        ax.axvspan(lower_limit_rect_start, lower_limit, color='red', alpha=0.3)
        ax.axvspan(upper_limit, upper_limit_rect_end, color='red', alpha=0.3)
    else:
        ax.axvspan(lower_limit_rect_start, lower_limit, color='red', alpha=0.3)
        ax.axvspan(upper_limit, upper_limit_rect_end, color='red', alpha=0.3)

    xlabel = get_analyte_label(analyte)
    sublabel = get_analyte_sublabel(analyte)
    ax.set_xlabel(f"{xlabel}\n{sublabel}", fontsize=12)
    ax.set_ylabel('Number of sites', fontsize=14)

    ax.set_xticks(x_ticks)
    ax.set_xticklabels([format_string.format(x) for x in x_ticks], fontsize=10)

    max_sites = max(n)
    if max_sites <= 5:
        y_ticks = np.arange(1, 6, 1)
    elif max_sites <= 10:
        y_ticks = np.arange(0, 12, 2)
    elif max_sites <= 30:
        y_ticks = np.arange(0, 36, 5)
    elif max_sites <= 50:
        y_ticks = np.arange(0, 60, 10)
    else:
        y_ticks = np.arange(0, max_sites + 10, 10)

    ax.set_yticks(y_ticks)
    ax.set_yticklabels([f'{int(y_tick)}' for y_tick in y_ticks], fontsize=12)
    ax.set_ylim(0, max(y_ticks))

    title_color = 'green' if interpretation == "Acceptable" else 'red'
    ax.set_title(f'{interpretation}', color=title_color, pad=20, fontsize=16, fontweight='bold')

    if analyte == "INR":
        ax.set_xlim(lower_limit_rect_start, upper_limit_rect_end) 
    else:
        ax.set_xlim(lower_limit_rect_start, upper_limit_rect_end)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig(img_stream, format='png')
    plt.close()
    img_stream.seek(0)
