# MacrosVBA

Macros del office para tareas frecuentes y repetitivas

## Importar las Macros

### Si te has descargado el archivo .bas
Para importar la macro, abre una instancia de Word, haz clic en “Archivo” y, a continuación, en “Importar archivo”. Tienes que especificar la ubicación del archivo de macros y hacer clic en “Abrir” para comenzar con la importación.
### Si sólo quieres importar el código
En el caso de que quieras importar sólo el código VBA de una macro, la importación se realiza de otra manera. Lo primero es seleccionar el documento al que se desea añadir la secuencia de comandos automatizada. Para ello, abre el Explorador de proyectos y haz doble clic en “Normal” (así la macro se guardará en la plantilla general) o en la entrada “ThisDocument” (subcarpeta de “Microsoft Word Objetos”).
Aparecerá una ventana de código que es donde vamos a copiar el código de la macro. Luego, hacemos clic en “Guardar”. Si en el paso anterior seleccionaste un documento de Word específico, se te informará de que es necesario guardarlo como “Documento de Word con macros”. Para ello, haz clic en “No” y selecciona la opción correcta bajo “Tipo de archivo”. Cuando hayas acabado, haz clic en “Guardar” para crear el nuevo formato de archivo.

## Descripción de las Macros:
### Draft:
Añade la marca de agua a todas las páginas y secciones del documento. De esta forma, estandarizamos el formato.

### Rellenator:
Buscará en todo el documento tablas en las cuales hayan celdas vacías y las rellenará con el texto que se haya introducido en el popup.