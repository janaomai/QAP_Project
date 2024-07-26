import pandas as pd

def preprocess_single_analyte(file_path, cycle, analyte):
    try:
        # Load the Excel file, assuming the data is in the first sheet
        df = pd.read_excel(file_path)

        # Ensure the cycle (column name) exists
        if cycle not in df.columns:
            return f"Error: Column '{cycle}' not found in the Excel file."

        # Skip the first row and handle specific cases
        column_data = df[cycle].iloc[1:]

        def handle_special_values(x):
            if isinstance(x, str) and x.startswith("<") and analyte == "Troponin":
                return 39
            if pd.isna(x) or isinstance(x, str):
                return "No submission"
            return x

        # Apply special values handling
        column_data = column_data.apply(handle_special_values)

        # Convert to numeric where possible, forcing errors to NaN
        numeric_data = pd.to_numeric(column_data, errors='coerce')

        # Calculate IQR and identify outliers
        Q1 = numeric_data.quantile(0.15)
        Q3 = numeric_data.quantile(0.85)
        IQR = Q3 - Q1
        lower_bound = Q1 - 1.5 * IQR
        upper_bound = Q3 + 1.5 * IQR

        def mark_outliers(x):
            if x == "No submission":
                return x
            if x < lower_bound or x > upper_bound:
                return "Unacceptable"
            return x

        column_data = column_data.apply(lambda x: mark_outliers(pd.to_numeric(x, errors='coerce')) if x != "No submission" else x)

        # Define the format mapping for different analytes
        format_mapping = {
            "D-Dimer": "{:.2f}",
            "CRP": "{:.1f}",
            "INR": "{:.1f}",
            "Troponin": "{:.0f}",
            "NT-ProBNP": "{:.0f}",
            "Haemoglobin": "{:.0f}",
            "Epoc": "{:.1f}",
            "i-STAT": "{:.1f}",
            "Lipids": "{:.1f}",
            "WCC": "{:.1f}"
        }

        # Apply the format transformation to the numeric data
        if analyte in format_mapping:
            format_string = format_mapping[analyte]
            if format_string == "int":
                column_data = column_data.apply(lambda x: int(x) if x not in ["No submission", "Unacceptable"] else x)
            else:
                column_data = column_data.apply(lambda x: format_string.format(float(x)) if x not in ["No submission", "Unacceptable"] else x)

        # Ensure all 'No submission' entries align with the "Site" column
        site_count = len(df["Site"].iloc[1:])
        if len(column_data) < site_count:
            column_data = column_data.append(pd.Series(["No submission"] * (site_count - len(column_data))), ignore_index=True)

        # Return the preprocessed data along with the Site column
        return df["Site"].iloc[1:], column_data
    except Exception as e:
        return f"Error: {str(e)}"

