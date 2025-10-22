import sqlite3

def get_patient(search_type, search_value, db_path="patients.db"):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    query = """
        SELECT first_name, last_name, dob, gender, patient_id, sample_id,
               analysis_id, sample_date, address, diagnosis
        FROM patients
    """
    if search_type == "Patient ID":
        query += " WHERE patient_id=?"
        cur.execute(query, (search_value,))
    elif search_type == "Sample ID":
        query += " WHERE sample_id=?"
        cur.execute(query, (search_value,))
    else:  # Name
        query += " WHERE last_name=? OR first_name=?"
        cur.execute(query, (search_value, search_value))

    result = cur.fetchone()
    conn.close()
    return result
