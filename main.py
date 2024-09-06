import os
from docx2pdf import convert


def docx_to_pdf(target_folder, pdf_folder):
    """
    DOCX fájlokat konvertál PDF formátumba és a PDF fájlokat a megadott PDF mappába menti.

    :param target_folder: A DOCX fájlok forrásmappája
    :param pdf_folder: A célmappa, ahova a PDF fájlok kerülnek
    """
    # Ellenőrizzük a PDF mappa létezését, ha már létezik, nem hozunk létre újat
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
        print(f"Létrehozva a 'PDFs' mappa: {pdf_folder}")
    else:
        print(f"A 'PDFs' mappa már létezik, az összes fájl újrakonvertálom.")

    # Végigmegyünk a forrásmappában lévő fájlokon
    for file_name in os.listdir(target_folder):
        if file_name.endswith(".docx"):
            docx_file_path = os.path.join(target_folder, file_name)
            pdf_file_path = os.path.join(pdf_folder, os.path.splitext(file_name)[0] + ".pdf")

            # DOCX konvertálása PDF-re
            convert(docx_file_path, pdf_file_path)
            print(f"Konvertálva: {docx_file_path} --> {pdf_file_path}")


# Felhasználótól bekérjük a DOCX fájlok mappáját
target_folder = input("Adja meg a DOCX fájlok mappáját: ")

# Létrehozzuk a 'PDFs' mappát a célmappán belül
pdf_folder = os.path.join(target_folder, "PDFs")

# DOCX fájlok konvertálása PDF-re
docx_to_pdf(target_folder, pdf_folder)

print("Az összes DOCX fájl sikeresen PDF formátumba lett konvertálva.")
