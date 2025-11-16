import csv

input_file = "regista_export_validat.csv"
output_file = "regista_export_curat_editat.csv"

# Lista corectă de headere, curățată
EXPECTED_HEADERS = [
    "Nr.", "Dată", "Nr. extern", "Emitent", "Conținut",
    "Compartimente", "Destinatar", "Activitate",
    "Stare doc.", "Tip", "Fișiere", "Operațiuni"
]

def normalize(row):
    """Îndepărtează spațiile, taburile, whitespace-ul și normalizează textul."""
    return [col.strip().replace("\ufeff", "") for col in row]

with open(input_file, "r", encoding="utf-8-sig", newline="") as fin, \
     open(output_file, "w", encoding="utf-8-sig", newline="") as fout:

    reader = csv.reader(fin)
    writer = csv.writer(fout)

    header_written = False

    for row in reader:
        clean = normalize(row)

        # sarim liniile complet goale
        if not any(clean):
            continue

        # detectăm header chiar dacă nu e perfect aliniat
        if clean[:len(EXPECTED_HEADERS)] == EXPECTED_HEADERS:
            if header_written:
                # acest header este repetat → îl ștergem
                continue
            else:
                # primul header → îl păstrăm
                header_written = True
                writer.writerow(EXPECTED_HEADERS)
                continue

        # orice alt rând se păstrează
        writer.writerow(row)

print("✔️ CSV curățat generat: regista_export_curat.csv (fără ștergeri greșite)")
