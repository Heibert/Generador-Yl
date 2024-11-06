"""This script generates a PDF file for each record in a excel"""

try:
    import os
    import re
    import sys
    import traceback

    import pandas as pd
    import pyfiglet

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

    # Check if the xlsx folder exists and create it if it doesn't
    if not os.path.exists(FILES_PATH + "xlsx"):
        os.makedirs(FILES_PATH + "xlsx")

    if not os.path.exists(FILES_PATH + "pdf"):
        os.makedirs(FILES_PATH + "pdf")

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
                    print(f"La columna '{column}' no se encuentra en el archivo '{file}'.")
                    input("Presiona enter para salir.")
                    sys.exit()
            return data
        except Exception as e:
            print(f"Hubo un error al leer el archivo {file}: {str(e)}")
            input("Presiona enter para salir.")
            sys.exit()

except Exception as e:
    print(traceback.format_exc())
    print(f"Hubo un error: {str(e)}")
    input("Presiona enter para salir.")
    sys.exit()
