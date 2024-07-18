import cv2
from pyzbar.pyzbar import decode
import pandas as pd
import openpyxl as xl
import io
import datetime

def verificarID(id_value, sheet):
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == id_value:
            return True
    return False

# Intentar abrir la cámara (puedes cambiar el índice 0 si es necesario)
cap = cv2.VideoCapture(0)

if not cap.isOpened():
    print("Error: No se puede abrir la cámara.")
    exit()

# Ruta del archivo de Excel
archivo_ruta = r'Evento.xlsx'

# Cargar el archivo de Excel y seleccionar la hoja
wb = xl.load_workbook(archivo_ruta)
registro_sheet = wb['registro']
checkIn_sheet = wb['checkIn']

while True:
    # Capturar el cuadro
    ret, frame = cap.read()

    # Verificar si el cuadro es válido
    if not ret or frame is None or frame.size == 0:
        print("Error: Cuadro de video inválido.")
        break
    
    for codes in decode(frame):
        # Decodificar el código QR 
        info = codes.data.decode('utf-8')
        
        # Verificar que la información no esté vacía
        if info:
                # Leer la cadena CSV desde la información decodificada
                data = pd.read_csv(io.StringIO(info))
                
                 # Convertir el contenido de registro a un DataFrame
                registro_df = pd.DataFrame(registro_sheet.values)
                registro_df.columns = registro_df.iloc[0]
                registro_df = registro_df[1:]
                
                # Verificar si el ID ya existe en la hoja de registro
                if verificarID(data['ID'][0], registro_sheet):
                    
                    # Verificar si el ID ya está en la hoja de checkIn
                    if verificarID(data['ID'][0], checkIn_sheet):
                        print(f"El ID {data['ID'][0]} ya ha realizado el checkIn.")
                    else :
                        # Agregar al DataFrame la columna asistencia con valor 1
                        data['Asistencia'] = 1
                        data ['Fecha'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        data = data[['ID', 'Nombre','Correo','Evento', 'Apellido', 'Asistencia','Fecha']]
                        
                        
                        # Agregar los nuevos datos a la hoja de checkIn
                        for index, row in data.iterrows():
                            checkIn_sheet.append(row.values.tolist())
                        print(f"Código QR guardado: \n {info} en la hoja de checkIn")
                else:
                    print(f"El ID {data['ID'][0]} no está registrado en la hoja de registro.")
            
                # Guardar el archivo Excel
                wb.save(archivo_ruta)

                    
    cv2.imshow("LECTOR DE QR", frame)
    
    # Salir si se presiona la tecla 'q'
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Liberar la cámara y cerrar las ventanas
cap.release()
cv2.destroyAllWindows()
