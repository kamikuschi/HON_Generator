from openpyxl import load_workbook

# Pfad zur Excel-Datei
file = './HON.xlsx'

# Excel-Datei laden
workbook = load_workbook(file)
sheet = workbook.active

# Listen für die Daten erstellen
days = []
duration = []
start_time = []
end_time = []

# Zeilen durchlaufen, beginnend mit der zweiten (die erste ist die Kopfzeile)
for row in sheet.iter_rows(min_row=5, min_col=2, max_col=5, values_only=True):
    if row[0] is not None:  # Überprüfen, ob Datum vorhanden ist
        days.append(row[0])
        start_time.append(row[2])
        end_time.append(row[3])
        duration.append(row[4])

# Ausgabe der extrahierten Daten
print("Tage:", days)
print("Dauer:", duration)
print("Anfangszeit:", start_time)
print("Endzeit:", end_time)
