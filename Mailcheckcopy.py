import pandas as pd

# Excel-Datei einlesen
df = pd.read_excel('path_to_file/input.xlsx')

# Längere und kürzere Spalte definieren
long_emails = df['Email_Long'].dropna().unique()
short_emails = df['Email_Short'].dropna().unique()

# Filtere die kürzere Spalte
filtered_short_emails = [email for email in short_emails if email not in long_emails]

# Ergebnis in einem neuen DataFrame speichern
result_df = pd.DataFrame(filtered_short_emails, columns=['Filtered_Email_Short'])

# Ergebnis in eine neue Excel-Datei speichern
result_df.to_excel('path_to_file/output.xlsx', index=False)

print("Gefilterte E-Mail-Adressen wurden in 'output.xlsx' gespeichert.")
