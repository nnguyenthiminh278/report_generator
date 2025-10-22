import sqlite3

# Create (or connect to) database
conn = sqlite3.connect("patients.db")
cur = conn.cursor()

# Create patients table
cur.execute("""
CREATE TABLE IF NOT EXISTS patients (
    id INTEGER PRIMARY KEY,
    name TEXT NOT NULL,
    dob TEXT NOT NULL,
    diagnosis TEXT
)
""")

# Insert sample patients
# Drop old table if exists (optional, for testing)
cur.execute("DROP TABLE IF EXISTS patients")

# Create patients table with more fields
cur.execute("""
CREATE TABLE patients (
    id INTEGER PRIMARY KEY,
    patient_id TEXT,
    sample_id TEXT,
    analysis_id TEXT,
    first_name TEXT,
    last_name TEXT,
    dob TEXT,
    gender TEXT,
    address TEXT,
    diagnosis TEXT,
    sample_date TEXT
)
""")

# Insert sample data
patients = [
    (
        1, "1234", "567890", "338252/53/54/56",
        "Alice", "Smith", "1985-04-12", "weiblich", "Berlin, Germany",
        "Chronic Kidney Disease", "2025-09-08"
    ),
    (
        2, "1235", "567891", "338252/53/54/57",
        "Bob", "Johnson", "1970-09-23", "männlich", "Munich, Germany",
        "Hypertension", "2025-09-08"
    ),
    (
        3, "1236", "567892", "338252/53/54/58",
        "Charlie", "Brown", "1992-01-15", "männlich", "Hamburg, Germany",
        "Diabetes Type II", "2025-09-08"
    )
]

cur.executemany("""
INSERT INTO patients (
    id, patient_id, sample_id, analysis_id, first_name, last_name, dob,
    gender, address, diagnosis, sample_date
) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
""", patients)

conn.commit()
conn.close()

print("patients.db created with extended sample data!")