import preprocsingle
import metricssingle
import zscore
import graphing
import reporting
import threading
import epoc_data_analysis
import istat_data_analysis
import lipids_data_analysis
import wbcdiff_data_analysis

file_path = ""
program_type = ""
analyte = ""
cycle_method = ""
cycle = ""
is_running_flag = threading.Event()

def set_file_path(path):
    global file_path
    file_path = path

def run_analysis_thread(program_type_value, analyte_value, cycle_method_value, cycle_value, user_id):
    try:
        run_analysis(program_type_value, analyte_value, cycle_method_value, cycle_value, user_id)
    except Exception as e:
        print(f"Error running analysis: {e}")
    finally:
        is_running_flag.set()  # Signal that the task is done

def run_analysis(program_type_value, analyte_value, cycle_method_value, cycle_value, user_id):
    global program_type, analyte, cycle_method, cycle, file_path
    program_type = program_type_value
    analyte = analyte_value
    cycle_method = cycle_method_value
    cycle = cycle_value

    if program_type == "Single Analyte":
        result = preprocsingle.preprocess_single_analyte(file_path, cycle, analyte)
        if isinstance(result, str) and result.startswith("Error"):
            print(result)
        else:
            sites, processed_data = result
            metrics_results, histograms = metricssingle.process_metrics(sites, processed_data, cycle, analyte)
            if metrics_results is None or histograms is None:
                print("Error processing metrics")
                return
            z_scores, no_submission_counter = zscore.calculate_z_scores(file_path, cycle)
            if isinstance(z_scores, str) and z_scores.startswith("Error"):
                print(z_scores)
            else:
                zscore.print_z_scores(z_scores, no_submission_counter)
                reporting.create_report(
                    template_path=r'C:\iCCnet QAP Program\Source_files\SingleTemplate.docx',
                    output_path=r"C:\iCCnet QAP Program\Output\POCT",
                    sites=sites,
                    metrics_data=metrics_results,
                    analyte=analyte,
                    cycle=cycle,
                    z_scores=z_scores,
                    no_submission_counter=no_submission_counter,
                    histograms=histograms,
                    file_path=file_path,
                    user_id=user_id
                )
    elif program_type == "Multi Analyte" and analyte == "Epoc":
        epoc_data_analysis.file_path = file_path
        epoc_data_analysis.sheet_name = cycle
        epoc_data_analysis.user_id = user_id
        epoc_data_analysis.run()

    elif program_type == "Multi Analyte" and analyte == "i-STAT":
        istat_data_analysis.file_path = file_path
        istat_data_analysis.sheet_name = cycle
        istat_data_analysis.user_id = user_id
        istat_data_analysis.run()

    elif program_type == "Multi Analyte" and analyte == "Lipids":
        lipids_data_analysis.file_path = file_path
        lipids_data_analysis.sheet_name = cycle
        lipids_data_analysis.user_id = user_id
        lipids_data_analysis.run()

    elif program_type == "Multi Analyte" and analyte == "WCC":
        wbcdiff_data_analysis.file_path = file_path
        wbcdiff_data_analysis.sheet_name = cycle
        wbcdiff_data_analysis.user_id = user_id
        wbcdiff_data_analysis.run()

def run_program_in_thread(program_type_val, analyte_val, cycle_method_val, cycle_val, user_id):
    is_running_flag.clear()  
    thread = threading.Thread(target=run_analysis_thread, args=(program_type_val, analyte_val, cycle_method_val, cycle_val, user_id))
    thread.start()
