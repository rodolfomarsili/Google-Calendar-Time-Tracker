```markdown
# Google Calendar Time Tracker

Este proyecto de Google Apps Script te permite rastrear tu tiempo dedicado a diferentes categorías de eventos en tu calendario de Google. 

## Funcionalidades:

- **Clasifica eventos:** Clasifica automáticamente los eventos del calendario en función de su color y una tabla de categorías personalizable.
- **Calcula horas trabajadas:** Calcula las horas dedicadas a cada categoría dentro de un rango de fechas especificado.
- **Exporta datos a Hoja de Cálculo:** Exporta los datos de seguimiento del tiempo a una hoja de cálculo de Google Sheets.
- **Interfaz de usuario amigable:** Proporciona una barra lateral en Google Sheets para configurar el seguimiento y actualizar los datos.

## Instalación:

1. **Crea una copia:** Haz una copia de esta carpeta en tu Google Drive.
2. **Abre la Hoja de Cálculo:** Abre el archivo `[Nombre del archivo de la hoja de cálculo].gsheet`.
3. **Autoriza el script:** La primera vez que ejecutes el script, deberás autorizarlo para que acceda a tu calendario y a la hoja de cálculo.
4. **Configura las categorías:** Accede a la tabla de colores desde la barra lateral para personalizar las categorías de eventos.

## Uso:

1. **Abre la barra lateral:** En la hoja de cálculo, ve a "Complementos" > "Google Calendar Time Tracker" > "Mostrar barra lateral".
2. **Selecciona el rango de fechas:** En la barra lateral, selecciona el rango de fechas para el que deseas rastrear tu tiempo.
3. **Actualiza la hoja de cálculo:** Haz clic en el botón "Actualizar Hoja" para procesar los datos de tu calendario y actualizar la hoja de cálculo.

## Personalización:

- **Tabla de colores:** Puedes personalizar las categorías de eventos y sus colores correspondientes en la tabla de colores accesible desde la barra lateral.
- **Hoja de cálculo de destino:** Puedes cambiar la hoja de cálculo y la hoja de destino modificando las variables `spreadsheetId` y `sheetName` en el archivo `[Nombre del archivo de código].gs`.

## Notas:

- Este script requiere acceso a tu calendario de Google y a Google Sheets.
- Asegúrate de que el evento "Working Hours" esté configurado correctamente en tu calendario para un seguimiento preciso del tiempo.

## Contribuciones:

Las contribuciones son bienvenidas. Por favor, crea un "fork" del repositorio y envía una solicitud de extracción con tus cambios.
``` 
