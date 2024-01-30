This Python script converts (Amtliches Liegenschaftskatasterinformationssystem) ALKIS information from Word documents (.docx) into a structured Excel format (.xlsx). It utilizes the docx, openpyxl, os, re, glob, and tkinter libraries. The code reads and analyzes specific details from the documents, such as event type, business reference, date, and personal information. It then creates an Excel file with organized data. The script includes a graphical user interface (GUI) using tkinter for user interaction, allowing users to choose between processing all files in a directory or selecting specific files.

Notes:

    The code is written in German, and comments and messages are in the same language.
    Certain data, such as variable names and sample information, has been modified to maintain privacy.
    Regular expressions are employed to extract specific details from each document.
    The use of tkinter provides a user-friendly interface for file selection and interaction.
    The code uses an object-oriented approach for structuring data analysis and manipulation.

 
script: 

import docx
from openpyxl import Workbook
import os
import re
import glob
import tkinter as tk
from tkinter import filedialog, messagebox

def lese_docx(datei):
    try:
        doc = docx.Document(datei)
        gesamter_text = [absatz.text for absatz in doc.paragraphs]
        return gesamter_text
    except Exception as e:
        print(f"Fehler beim Lesen der Datei {datei}: {e}")
        return None

def analysiere_text(text, dateiname):
    daten = {
        'Datei': dateiname,
        'Anlassart': '',
        'Geschäftszeichen': '',
        'Datum': '',
        'Anteil': '',
        'Zuname': '',
        'Vorname': '',
        'Geburtsname': '',
        'Geburtsdatum': '',
        'Strasse_Hsnr': '',
        'PLZ': '',
        'Ort': ''
    }

    for zeile in text:
        if 'F   00 SOL2' in zeile:
            match = re.search(r'F   00 SOL2\s+(\d+)', zeile)
            if match:
                daten['Anlassart'] = match.group(1)
        elif 'GZ:' in zeile:
            match = re.search(r'GZ:\s+(\S+),', zeile)
            if match:
                daten['Geschäftszeichen'] = match.group(1)
            match = re.search(r'DATUM:\s+(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})', zeile)
            if match:
                daten['Datum'] = match.group(1)
        elif '1  0' in zeile:
            match = re.search(r'1\s+0\s+(.*)', zeile)
            if match:
                daten['Anteil'] = match.group(1)
        elif '1  1' in zeile:
            match = re.search(r'1\s+1(\w+),', zeile)
            if match:
                daten['Zuname'] = match.group(1)
            match = re.search(r'1\s+1\w+,\s*(.*),,,', zeile)
            if match:
                daten['Vorname'] = match.group(1)
        elif '1  2' in zeile:
            match = re.search(r'1\s+2+geb\.\s+(\w+)', zeile)
            if match:
                daten['Geburtsname'] = match.group(1)
            match = re.search(r'\*([\d\.]+)', zeile)
            if match:
                daten['Geburtsdatum'] = match.group(1)
        elif '1  3' in zeile:
            match = re.search(r'1\s+3(.*)', zeile)
            if match:
                daten['Strasse_Hsnr'] = match.group(1)
        elif '1  4' in zeile:
            match = re.search(r'1\s+4(\d{5})\s+(.*)', zeile)
            if match:
                daten['PLZ'] = match.group(1)
                daten['Ort'] = match.group(2)

    return [daten]


def erstelle_excel(daten, excel_dateipfad):
    try:
        wb = Workbook()
        ws = wb.active
        spaltenüberschriften = ['Datei', 'Anlassart', 'Geschäftszeichen', 'Datum', 'Anteil', 'Zuname', 'Vorname', 'Geburtsname', 'Geburtsdatum', 'Strasse_Hsnr', 'PLZ', 'Ort']
        ws.append(spaltenüberschriften)

        for datenblock in daten:
            ws.append([datenblock.get(spalte, '') for spalte in spaltenüberschriften])

        wb.save(excel_dateipfad)
    except Exception as e:
        print(f"Fehler beim Erstellen der Excel-Datei {excel_dateipfad}: {e}")

def verarbeite_alle_dateien(verzeichnispfad):
    dateipfade = glob.glob(os.path.join(verzeichnispfad, '*.docx'))

    for dateipfad in dateipfade:
        text = lese_docx(dateipfad)
        if text:
            dateiname = os.path.basename(dateipfad)
            strukturierte_daten = analysiere_text(text, dateiname)
            excel_dateipfad = os.path.splitext(dateipfad)[0] + '.xlsx'
            erstelle_excel(strukturierte_daten, excel_dateipfad)
            print(f"Konvertierung abgeschlossen für: {excel_dateipfad}")

def verarbeite_ausgewählte_dateien():
    dateipfade = filedialog.askopenfilenames(
        title='Wählen Sie die zu konvertierenden Dateien aus',
        filetypes=[('Word-Dokumente', '*.docx')]
    )

    for dateipfad in dateipfade:
        text = lese_docx(dateipfad)
        if text:
            dateiname = os.path.basename(dateipfad)
            strukturierte_daten = analysiere_text(text, dateiname)
            excel_dateipfad = os.path.splitext(dateipfad)[0] + '.xlsx'
            erstelle_excel(strukturierte_daten, excel_dateipfad)
            print(f"Konvertierung abgeschlossen für: {excel_dateipfad}")

def main():
    root = tk.Tk()
    root.withdraw()

    auswahl = messagebox.askyesno("Auswahl", "Möchten Sie alle Dateien in einem Verzeichnis konvertieren?")
    if auswahl:
        verzeichnispfad = filedialog.askdirectory()
        if verzeichnispfad:
            verarbeite_alle_dateien(verzeichnispfad)
    else:
        verarbeite_ausgewählte_dateien()

if __name__ == "__main__":
    main()

Author: Andersson Barbosa
Note: This code was developed to automate the weekly conversion process when new files arrive and are adapted to readable formats. 
