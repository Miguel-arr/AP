import win32com.client as win32
import re
import pandas as pd

def codsia_outlook():
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 is the index of the inbox
    codigos = []
    datos = []

    while True:
        messages = inbox.Items.Restrict("[Unread] = true AND [Subject] = 'Image processing complete'")

        if messages.Count == 0:
            break
        print(f"Total de correos no leídos con el asunto 'Image processing complete': {messages.Count}")

        for message in messages:
            try:
                # Buscar el código en el cuerpo del correo
                match_code = re.search(r'Processing completed for (\w+)', message.Body)  # \w+ coincide con cualquier combinación de letras y números
                match_flight_date = re.search(r'Flight Date: (\d{2}/\d{2}/\d{4})', message.Body)  # Busca la fecha del vuelo en formato DD/MM/YYYY
                if match_code and match_flight_date:
                    codigo = match_code.group(1)
                    flight_date = match_flight_date.group(1)
                    received_date = message.ReceivedTime.strftime('%d/%m/%Y')
                    print("Código encontrado:", codigo)
                    print("Fecha de vuelo encontrada:", flight_date)
                    print("Fecha de recibo del correo:", received_date)
                    codigos.append(codigo)
                    datos.append({
                        'Codigo': codigo,
                        'Fecha Correo': received_date,
                        'Fecha Vuelo': flight_date
                    })
                    # Marcar como leído
                    message.UnRead = False
            except Exception as e:
                print(f'Error al procesar el mensaje: {e}')
    
    # Crear un DataFrame de pandas con los datos recolectados
    df = pd.DataFrame(datos)
    
    ruta_guardado = r'C:\Users\marodriguezr\Documents\fechas\codigos_y_fechas.xlsx'
    df.to_excel(ruta_guardado, index=False)  # Aquí se exporta el archivo Excel
    print(f"Datos exportados a {ruta_guardado}")

    return codigos

# Llamar a la función para ejecutar el script
codigos = codsia_outlook()

# codigos contiene solo los códigos
print(codigos)
