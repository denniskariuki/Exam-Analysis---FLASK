
import os
import pandas as pd
import openpyxl
import re

def process_and_analyze(file_path, exam_type, output_folder):
    """
    Cleans and analyzes an Excel file for exam results.
    :param file_path: Path to the uploaded Excel file
    :param exam_type: 'CDACC' or 'KNEC'
    :param output_folder: Folder to save analyzed files
    :return: Path to the analyzed report
    """
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Load the Excel file
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Extract file name from 'Class :' row
    file_name = None
    row_to_delete = None
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "Class :" in cell.value:
                file_name = cell.value.split(":")[1].strip()
                row_to_delete = row_idx
                break
        if file_name:
            break

    if not file_name:
        raise ValueError("Class name not found in the file")

    # Delete the extracted row and first four rows
    if row_to_delete:
        ws.delete_rows(row_to_delete)
    ws.delete_rows(1, 4)

    # Remove images if present
    if ws._images:
        ws._images.clear()

    # Convert sheet to DataFrame
    data = pd.DataFrame(ws.values)
    data.columns = data.iloc[0]
    data = data[1:]

    # Remove unnecessary columns
    columns_to_remove = ["Result", "Classification"]
    data = data.drop(columns=[col for col in columns_to_remove if col in data.columns], errors='ignore')

    # Remove rows with specific labels
    rows_to_remove = [
        "Total Students", "Lowest Score", "Highest Score",
        "Total Marks", "Average Marks", "Average Grade", "Average Competence"
    ]
    data = data[~data.iloc[:, 0].astype(str).isin(rows_to_remove)]

    # Remove columns containing '#'
    columns_to_remove = [col for col in data.columns if "#" in str(col)]
    data = data.drop(columns=columns_to_remove, errors='ignore')

    # Save cleaned file
    cleaned_file = os.path.join(output_folder, f"{file_name}_cleaned.xlsx")
    data.to_excel(cleaned_file, index=False)

    # --- Analysis Stage ---
    df = pd.read_excel(cleaned_file)
    df.fillna("", inplace=True)
    df.columns = df.columns.str.strip()
    df.rename(columns={"Admission No": "Admission Number"}, inplace=True)

    # Identify subjects
    required_columns = {"Admission Number", "Student Name"}
    subject_columns = [col for col in df.columns if col not in required_columns]

    students_data, missing_marks_data, failed_units_data = [], [], []
    for _, row in df.iterrows():
        admission_number = row.get("Admission Number", "Unknown")
        student_name = row.get("Student Name", "Unknown")
        total_score, valid_subjects_count, student_marks = 0, 0, []
        missed_subjects, failed_subjects = [], []

        for subject in subject_columns:
            value = str(row[subject]).strip()
            if "MM" in value:
                missed_subjects.append(subject)
            match = re.search(r"\d+", value)
            if match:
                score = float(match.group())
                total_score += score
                valid_subjects_count += 1
                student_marks.append(score)
                if (exam_type == "CDACC" and score < 50) or (exam_type == "KNEC" and score < 40):
                    failed_subjects.append(subject)

        avg_score = round(total_score / valid_subjects_count, 2) if valid_subjects_count > 0 else 0
        
        # Determine competency level
        if exam_type == "CDACC":
            if avg_score < 50:
                competency = "Not Yet Competent"
            elif 50 <= avg_score < 70:
                competency = "Competent"
            elif 70 <= avg_score < 85:
                competency = "Proficient"
            else:
                competency = "Mastery"
        elif exam_type == "KNEC":
            if avg_score >= 80:
                competency = "Distinction 1"
            elif 75 <= avg_score < 80:
                competency = "Distinction 2"
            elif 70 <= avg_score < 75:
                competency = "Credit 3"
            elif 60 <= avg_score < 70:
                competency = "Credit 4"
            elif 50 <= avg_score < 60:
                competency = "Pass 5"
            elif 40 <= avg_score < 50:
                competency = "Pass 6"
            else:
                competency = "Fail"

      
        # Create a dictionary to store student details
        student_record = {
            "Admission Number": admission_number,
            "Student Name": student_name,
            "Average Score": avg_score,
            "Competency Level": competency
        }

        # Add each subject's score as a separate column
        for subject in subject_columns:
            match = re.search(r"\d+", str(row[subject]))
            student_record[subject] = float(match.group()) if match else ""  # Leave blank if no mark

        students_data.append(student_record)  # Append the updated dictionary

        for subject in missed_subjects:
            missing_marks_data.append({"Admission Number": admission_number, "Student Name": student_name, "Missed Subject": subject})
        for subject in failed_subjects:
            failed_units_data.append({"Admission Number": admission_number, "Student Name": student_name, "Referred Subject": subject})

    # Save analyzed report
    report_file = os.path.join(output_folder, f"{file_name}.xlsx")
    with pd.ExcelWriter(report_file) as writer:
        pd.DataFrame(students_data).to_excel(writer, sheet_name="Student Performance", index=False)
        pd.DataFrame(missing_marks_data).to_excel(writer, sheet_name="Missing Marks", index=False)
        pd.DataFrame(failed_units_data).to_excel(writer, sheet_name="Referred Units", index=False)

    return report_file

