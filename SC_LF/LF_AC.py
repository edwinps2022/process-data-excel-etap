import pandas as pd
import os

def procesar_reporte_etap(ruta_archivo_excel):
    """
    Procesa un reporte de ETAP en formato Excel para extraer y analizar los datos.

    :param ruta_archivo_excel: La ruta al archivo de Excel a procesar.
    """
    try:
        # Extraer el nombre base del archivo para usarlo en los nombres de los archivos de salida
        nombre_base = os.path.splitext(os.path.basename(ruta_archivo_excel))[0]

        # Leer el archivo de Excel
        # Es posible que necesites ajustar el nombre de la hoja 'sheet_name'
        df = pd.read_excel(ruta_archivo_excel, sheet_name='AC Load Flow Report')

        # ----- Aquí es donde agregaremos el código para procesar los datos -----
        # Por ejemplo, podrías querer:
        # 1. Limpiar los datos (eliminar filas/columnas innecesarias)
        # 2. Filtrar datos (por ejemplo, solo buses con bajo voltaje)
        # 3. Realizar cálculos (por ejemplo, promedios, máximos, etc.)
        # 4. Generar un nuevo archivo de Excel con los resultados

        print("El archivo de Excel se ha leído correctamente. El DataFrame tiene {df.shape[0]} filas y {df.shape[1]} columnas.")

        # Ejemplo de cómo guardar los datos procesados en un nuevo archivo de Excel
        # ruta_salida = f"{nombre_base}_procesado.xlsx"
        # df.to_excel(ruta_salida, index=False)
        # print(f"Archivo procesado guardado en: {ruta_salida}")

    except FileNotFoundError:
        print(f"Error: El archivo no se encontró en la ruta especificada: {ruta_archivo_excel}")
    except Exception as e:
        print(f"Ocurrió un error al procesar el archivo: {e}")

if __name__ == "__main__":
    # Pídele al usuario que ingrese la ruta del archivo de Excel
    ruta_del_archivo = input("Por favor, introduce la ruta completa al archivo de Excel de ETAP: ")
    
    # Llama a la función para procesar el reporte
    procesar_reporte_etap(ruta_del_archivo)
