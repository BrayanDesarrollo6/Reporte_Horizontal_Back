import pandas as pd

def procesar_excel(ruta_archivo: str):
    try:
        # Leer el archivo Excel
        df = pd.read_excel(ruta_archivo, usecols="A,B")

        # Obtener y validar los valores de la columna A (Valores)
        valores_columna_a = df.iloc[:, 0].dropna().tolist()
        valores_filtrados = [valor for valor in valores_columna_a if isinstance(valor, (int, float)) and valor > 0]

        if not valores_filtrados:
            raise ValueError("Error no se encontraron valores validos en la columna A")

        # Obtener y validar el valor de la columna B (Objetivo)
        objetivo_columna_b = df.iloc[:, 1].dropna().iloc[0]

        if not isinstance(objetivo_columna_b, (int, float)) or objetivo_columna_b <= 0:
            raise ValueError("Error el valor en la columna B no es numerico o no es mayor a 0")

        return valores_filtrados, objetivo_columna_b

    except FileNotFoundError:
        print(f"Error: El archivo en la ruta '{ruta_archivo}' no se encontro")
    except ValueError as error:
        print(f"Error de valor: {error}")
    except Exception as error:
        print(f"Ocurrio un error inesperado: {error}")
        return None, None