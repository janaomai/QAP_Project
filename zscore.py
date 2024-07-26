import pandas as pd

def calculate_z_scores(file_path, cycle):
    try:
        # Load the Excel file
        df = pd.read_excel(file_path)

        # Find all EQA columns
        eqa_columns = [col for col in df.columns if col.startswith("EQA")]

        # Ensure the selected cycle exists
        if cycle not in eqa_columns:
            return f"Error: Cycle '{cycle}' not found in the Excel file."

        # Get the index of the selected cycle
        cycle_index = eqa_columns.index(cycle)

        # Extract data for all previous EQA cycles
        previous_eqa_columns = eqa_columns[:cycle_index+1]
        data = df[previous_eqa_columns]

        # Initialise the result dictionary
        z_scores = {site: {} for site in df["Site"].iloc[1:]}

        # Initialise the no submission counter
        no_submission_counter = {site: 0 for site in df["Site"].iloc[1:]}

        # Calculate the Z-scores for each site
        for site, row in data.iloc[1:].iterrows():
            for col in previous_eqa_columns:
                value = row[col]
                if isinstance(value, str) or pd.isna(value):
                    z_scores[df["Site"].iloc[site]][col] = "No submission"
                    no_submission_counter[df["Site"].iloc[site]] += 1
                else:
                    # Reset counter if there is a submission
                    no_submission_counter[df["Site"].iloc[site]] = 0

                    # Calculate mean and standard deviation for the column excluding nan
                    col_values = pd.to_numeric(data[col].iloc[1:], errors='coerce').dropna()
                    if not col_values.empty:
                        mean = col_values.mean()
                        std = col_values.std()
                        z_score = (value - mean) / std if std != 0 else 0
                        z_scores[df["Site"].iloc[site]][col] = z_score
                    else:
                        z_scores[df["Site"].iloc[site]][col] = "No submission"

        return z_scores, no_submission_counter
    except Exception as e:
        return f"Error: {str(e)}", {}

def print_z_scores(z_scores, no_submission_counter):
    try:
        for site, scores in z_scores.items():
            acceptable_count = 0
            warning_count = 0
            unacceptable_count = 0
            for col, score in scores.items():
                if score == "No submission":
                    continue
                if isinstance(score, (int, float)):
                    if -1 <= score <= 1:
                        acceptable_count += 1
                    elif -2 <= score < -1 or 1 < score <= 2:
                        warning_count += 1
                    elif score <= -3 or score >= 3:
                        unacceptable_count += 1
            no_submission_str = ""
            if no_submission_counter[site] >= 2:
                no_submission_str = f"Consecutive No Submissions: {no_submission_counter[site]} - NON COMPLIANCE"
            elif no_submission_counter[site] > 0:
                no_submission_str = f"Consecutive No Submissions: {no_submission_counter[site]}"

            print(f"Generating graphs for {site}")
            if no_submission_str:
                print(f"Report generated.")
            else:
                print(f"Report generated.")
    except Exception as e:
        print(f"Error printing Z-scores: {str(e)}")

