from datetime import datetime

def build_context(patient_data):
    (first_name, last_name, dob, gender,
     patient_id_str, sample_id, analysis_id,
     sample_date, address, diagnosis) = patient_data

    anrede = "Frau" if str(gender).lower().strip() in ["weiblich", "female", "f"] else "Herr"
    report_date = datetime.now().strftime("%d.%m.%Y")

    return {
        "anrede": anrede,
        "vorname": first_name,
        "name": last_name,
        "dob": dob,
        "geschlecht": gender,
        "patient_id": patient_id_str,
        "sample_id": sample_id,
        "analysis_id": analysis_id,
        "sample_date": sample_date,
        "address": address,
        "diagnosis": diagnosis,
        "report_date": report_date,
        "patient_sample": f"{patient_id_str}-{sample_id}"
    }
