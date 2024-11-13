"""This script generates a PDF file for each record in a excel"""

try:
    import os
    import re
    import sys
    import traceback

    import pandas as pd
    import pyfiglet
    from docx import Document
    from docx2pdf import convert

    FILES_PATH = os.path.dirname(os.path.abspath(__file__)) + "\\"

    def disclaimer():
        ascii_art = pyfiglet.figlet_format("C & C", font="doh")

        lines = ascii_art.splitlines()

        # Find the first non-empty line from the top
        start_line = 0
        for i, line in enumerate(lines):
            if line.strip():
                start_line = i
                break

        # Find the last non-empty line from the bottom
        end_line = len(lines) - 1
        for i in range(len(lines) - 1, -1, -1):
            if lines[i].strip():
                end_line = i
                break

        # Crop the ASCII art to remove empty lines
        ASCCI = "\n".join(lines[start_line : end_line + 1])

        print(ASCCI)
        print()

        print(
            "\nRecuerde que este archivo es privado y no debe ser compartido con personal no autorizado.\n"
        )

    # Check if the folders exists and create them if they don't
    if not os.path.exists(FILES_PATH + "words"):
        os.makedirs(FILES_PATH + "words")

    if not os.path.exists(FILES_PATH + "pdf"):
        os.makedirs(FILES_PATH + "pdf")

    if not os.path.exists(FILES_PATH + "Plantillas"):
        print("No se encontró la carpeta 'Plantillas'.")
        input("Presiona enter para salir.")
        sys.exit()

    if not os.path.exists(FILES_PATH + "BASE BOT CORRESPONDENCIA.xlsx"):
        print("No se encontró el archivo 'BASE BOT CORRESPONDENCIA.xlsx'.")
        input("Presiona enter para salir.")
        sys.exit()

    def get_numbers(string):
        """Get the numbers from a string."""
        number = re.sub(r"[^0-9,]", "", string)
        if number == "":
            print(f"No se encontraron números en '{string}'.")
            return 0
        else:
            number = number.replace(",", ".")
            return float(number)

    def get_data_from_excel(file):
        """Get the data from the excel file."""
        try:
            # Read the data from the excel file
            data = pd.read_excel(file)
            # Ensure that the columns are in the correct order
            required_columns = [
                "CEDULA",
                "CODIGO",
                "DIRECTORA",
                "NOMBRE DE DIRECTORA",
                "CORREO DIR",
                "NOMBRE CONSULTORA",
                "DIRECCION",
                "BARRIO",
                "CIUDAD",
                "FACTURA",
                "VENCIMIENTO",
                "VALOR TOTAL AL DIA",
                "DIAS MORA AL DIA",
                "EDAD DE LIQUIDACION",
            ]
            for column in required_columns:
                if column not in data.columns:
                    print(
                        f"La columna '{column}' no se encuentra en el archivo '{file}'."
                    )
                    input("Presiona enter para salir.")
                    sys.exit()
            return data
        except Exception as e:
            print(f"Hubo un error al leer el archivo {file}: {str(e)}")
            input("Presiona enter para salir.")
            sys.exit()

    # Function to load and replace placeholders in the .docx
    def load_and_replace_docx(docx_path, replacements):
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            for placeholder, new_value in replacements.items():
                if placeholder in paragraph.text:
                    # Replace placeholder with actual value
                    paragraph.text = re.sub(placeholder, new_value, paragraph.text)

        # Save modified document to a temporary file
        modified_docx_path = "/mnt/data/Modified_Template.docx"
        doc.save(modified_docx_path)
        return modified_docx_path

    # Paths to the files
    original_docx_path = FILES_PATH + "/Plantillas/Plantilla 10-30.docx"
    output_pdf_path = FILES_PATH + "/pdf/Invoice.pdf"

    # Dictionary with placeholders and their corresponding values
    replacements = {
        "XDATE_TODAYX": "2024-11-13",
        "XCONSULTANT_NAMEX": "John Doe",
        "XADDRESSX": "123 Main Street",
        "XCITYX": "Bogotá",
        "XBILLX": "456789",
        "XEXPIRATION_DATEX": "2024-12-01",
        "XVALUEX": "1,000,000 COP",
    }

    # Load, replace placeholders, and save as a new .docx file
    modified_docx_path = load_and_replace_docx(original_docx_path, replacements)

    # Convert the modified .docx to PDF
    convert(modified_docx_path, output_pdf_path)

    print("PDF created successfully at:", output_pdf_path)

    data = get_data_from_excel(FILES_PATH + "BASE BOT CORRESPONDENCIA.xlsx")
    print(data)
    print(data['CEDULA'][0])
    # input("Presiona enter para salir.")

except Exception as e:
    print(traceback.format_exc())
    print(f"Hubo un error: {str(e)}")
    input("Presiona enter para salir.")
    sys.exit()
