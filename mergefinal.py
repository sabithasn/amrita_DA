# the program requires the name of the campus code as the last two letters 
# the school names should be the second last word in the file name. All in capital letters.
# Remember to adjust the school names to match the criterion of the program. 
# for CB, ASE in Phd and MSc, will have differnt school names, Adjust file names accordingly or do it manually
# This program formats the output to match the format in which the master Db is stored.
# It creates two files for parents and students which need to be updated to the end of the respective master DB
#********************************************************************************
#
#
# FINAL UPDTED ONE REQUIRED FORMAT
#
#********************************************************************************
import pandas as pd
import os
import re

# 🎓 Campus and School decoding dictionaries
campus_codes = {
    'AM': 'Amritapuri', 'CH': 'Chennai', 'MY': 'Mysore', 'BN': 'Bangalore',
    'KH': 'Kochi', 'CB': 'Coimbatore', 'AV': 'Amaravati', 'AH': 'Ahead'
}

school_codes = {
    'ASAS': 'Amrita School of Arts and Sciences', 'ASB': 'Amrita School of Business',
    'ASC': 'Amrita School of Computing', 'BT': 'Amrita School of Life Sciences',
    'ASCOM': 'Amrita School of Mass Communication', 'ASE': 'Amrita School of Engineering',
    'ASPS': 'Amrita School of Physical Sciences', 'ASBS': 'Amrita School of Behavioral Sciences',
    'AG': 'Amrita School of Agriculture', 'AST': 'Amrita School of Inter Disciplinary Studies',
    'CSS': 'Amritadarshanam', 'AY': 'Ayurveda', 'Ahead': 'Amrita Ahead/ Online'
}

def decode_filename(filename):
    name = os.path.splitext(filename)[0]
    parts = re.split(r'[_\s\-]+', name)

    school = None
    campus = None

    for part in parts:
        if part in school_codes:
            school = school_codes[part]
        if part in campus_codes:
            campus = campus_codes[part]

    return school or 'Unknown School', campus or 'Unknown Campus'

def extract_parent_info(file_path, school, campus):
    source_name = os.path.basename(file_path)

    try:
        df = pd.read_excel(file_path, sheet_name="Contact Details", skiprows=2, header=None)
        df.columns = df.iloc[0].map(str).str.strip()
        df = df.iloc[1:].reset_index(drop=True)

        selected = pd.DataFrame()
        selected["Roll Number"] = df.iloc[:, 0]
        selected["Father Name"] = df.iloc[:, 1]
        selected["Father Phone"] = df.iloc[:, 6]
        selected["Father Mobile"] = df.iloc[:, 7]
        selected["Father Personal Email"] = df.iloc[:, 8]
        selected["Father Work Email"] = df.iloc[:, 9]
        selected["Mother Name"] = df.iloc[:, 11]
        selected["Mother Phone"] = df.iloc[:, 16]
        selected["Mother Mobile"] = df.iloc[:, 17]
        selected["Mother Personal Email"] = df.iloc[:, 18]
        selected["Mother Work Email"] = df.iloc[:, 19]
        selected["Source File"] = source_name
        selected["School"] = school
        selected["Campus"] = campus

        return selected

    except Exception as e:
        print(f"⚠️ Skipping {source_name}: {e}")
        return pd.DataFrame()

def process_student_info(file_path, school, campus):
    source_name = os.path.basename(file_path)

    try:
        xls = pd.ExcelFile(file_path)
        required_sheets = ['General Details', 'Personal Details', 'Address Details']
        if not all(sheet in xls.sheet_names for sheet in required_sheets):
            print(f"⚠️ Skipping {source_name} — missing required student sheets.")
            return []

        general_df = pd.read_excel(xls, sheet_name='General Details', skiprows=2)
        personal_df = pd.read_excel(xls, sheet_name='Personal Details', skiprows=2)
        address_df = pd.read_excel(xls, sheet_name='Address Details', skiprows=2)

        merged_df = general_df.merge(personal_df, on="Roll Number", how="inner") \
                              .merge(address_df, on="Roll Number", how="inner")

        student_data = []
        for _, row in merged_df.iterrows():
            roll = str(row.get("Roll Number", ""))
            join_year = row.get("Joining Year", 0)

            # Extract program duration from 8th character of roll number
            try:
                duration = int(roll[7]) if roll and roll[7].isdigit() else 0
                if duration > 6:
                    duration = 0
            except:
                duration = 0

            grad_year = join_year + duration if pd.notna(join_year) else None

            address_parts = [
                str(row.get("Adress1", "")), str(row.get("Adress2", "")),
                str(row.get("City", "")), str(row.get("District", "")),
                str(row.get("State", "")), str(row.get("Country", "")),
                str(row.get("Pincode", ""))
            ]
            full_address = " ".join([part for part in address_parts if pd.notna(part)])

            student_row = {
                "Roll Number": roll,
                "Student Name": row.get("Student Name"),
                "Academic Program": row.get("Academic Program"),
                "Branch": row.get("Branch"),
                "Joining Year": join_year,
                "Gender": row.get("Gender"),
                "Date Of Birth": row.get("Date Of Birth"),
                "Date Of Enrollment": row.get("Date Of Enrollment"),
                "Remarks": row.get("Remarks"),
                "Year of Graduation": grad_year,
                "Campus": campus,
                "Nationality": row.get("Nationality"),
                "Native Place": row.get("Native Place"),
                "Email": row.get("Email"),
                "PhoneNo": row.get("PhoneNo"),
                "Address": full_address,
                "District": row.get("Native District"),
                "State": row.get("Native State"),
                "Country": row.get("Native Country"),
                "Program Duration": duration,
                "School": school,
                "Source File": source_name
                
            }
            student_data.append(student_row)

        return student_data

    except Exception as e:
        print(f"🚫 Error processing student info in {source_name}: {e}")
        return []

def merge_all_excels(folder_path):
    all_students = []
    all_parents = []

    for filename in os.listdir(folder_path):
        if filename.endswith(".xls") or filename.endswith(".xlsx"):
            full_path = os.path.join(folder_path, filename)
            school, campus = decode_filename(filename)

            student_data = process_student_info(full_path, school, campus)
            if student_data:
                all_students.extend(student_data)

            parent_df = extract_parent_info(full_path, school, campus)
            if not parent_df.empty:
                all_parents.append(parent_df)

    if all_students:
        student_df = pd.DataFrame(all_students)
        student_columns = [
            "Roll Number", "Student Name", "Academic Program", "Branch", "Joining Year",
            "Gender", "Date Of Birth", "Date Of Enrollment", "Remarks", "Year of Graduation",
            "Campus", "Nationality", "Native Place", "Email", "PhoneNo", "Address",
            "District", "State", "Country", "Program Duration","School", "Source File"
        ]
        student_df = student_df[student_columns]
        student_df.to_excel("student_details.xlsx", index=False)
        print("✅ student_details.xlsx created.")

    if all_parents:
        parent_df = pd.concat(all_parents, ignore_index=True)
        parent_columns = [
            "Roll Number", "Father Name", "Father Phone", "Father Mobile",
            "Father Personal Email", "Father Work Email",
            "Mother Name", "Mother Phone", "Mother Mobile",
            "Mother Personal Email", "Mother Work Email",
            "Source File", "School", "Campus"
        ]
        parent_df = parent_df[parent_columns]
        parent_df.to_excel("parents_details.xlsx", index=False)
        print("✅ parents_details.xlsx created.")

# 🏁 Run the merger

merge_all_excels(".")
