import sys
import os


carpeta = sys.argv[1]

# Eliminar todos los archivos en la carpeta
for archivo in os.listdir(carpeta):
    ruta_archivo = os.path.join(carpeta, archivo)
    if os.path.isfile(ruta_archivo):
        os.remove(ruta_archivo)

print('Archivos eliminados correctamente.')