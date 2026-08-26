import pandas as pd

EXPECTED_COLUMNS = [
    'Postcode', 'Age', 'Date Of Birth', 'Gender', 'Ethnicity',
    'MRC Code', 'Date', 'COPD Diagnosis', 'Date', 'Depression', 'Date2',
    'Anxiety', 'Date3', 'Referral to Pulmonary Rehab', 'Date4',
    'Smoking Status', 'Date5', 'Associated Text', 'Vaccinations', 'Date6',
    'Review of Inhaler Technique', 'Date7', 'Name, Dosage and Quantity',
    'Date of Issue2', 'Dose', 'Date of Issue', 'Name, Dosage and Quantity2',
    'Dose2', 'Date of Issue3', 'Name, Dosage and Quantity3', 'Dose3',
    'Consultations coded COPD', 'Date8'
]

df = pd.read_csv('Doccla_Evaluation_Report_09.07.csv')
df.columns = df.columns.str.strip()

assert list(df.columns) == EXPECTED_COLUMNS, "Column mismatch - stop and check."

print("=== SHAPE ===")
print(df.shape)

print("\n=== NULL COUNT PER COLUMN ===")
print(df.isna().sum())

print("\n=== CANDIDATE PATIENT-ANCHOR COLUMNS: agreement check ===")
# Does every row where Age is filled also have Postcode and DOB filled?
# And vice versa - do these three columns always go null/non-null together?
anchors = ['Age', 'Postcode', 'Date Of Birth']
notna = df[anchors].notna()
print(notna.value_counts())  # combinations of which anchors are filled together

print("\n=== ROW TYPE BREAKDOWN ===")
# Does having Age filled always coincide with row 0 of a new patient block?
has_age = df['Age'].notna()
has_inhaler = df['Name, Dosage and Quantity'].notna()
has_consultation = df['Consultations coded COPD'].notna()
print("Rows with Age filled:", has_age.sum())
print("Rows with inhaler data filled:", has_inhaler.sum())
print("Rows with consultation data filled:", has_consultation.sum())
print("Rows with BOTH inhaler and consultation filled:", (has_inhaler & has_consultation).sum())
print("Rows with NEITHER Age, inhaler, nor consultation filled:",
      (~has_age & ~has_inhaler & ~has_consultation).sum())

print("\n=== SAMPLE PATTERN: first 30 rows, presence-only (no values) ===")
presence = df.notna().astype(int)
presence.columns = [c[:15] for c in presence.columns]  # shorten for display
print(presence.head(30).to_string())

print("\n=== DUPLICATE 'Date' COLUMN CHECK ===")
# how many columns are literally named 'Date' before renaming
raw_cols = pd.read_csv('Doccla_Evaluation_Report_09.07.csv', nrows=0).columns
raw_cols_stripped = raw_cols.str.strip()
print([c for c in raw_cols_stripped if list(raw_cols_stripped).count(c) > 1])