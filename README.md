# PySSforM
Esta aplicacion creará automaticamente un formulario mediante un archivo de hojas de calculo.
Para hacer un buen uso de la aplicación hay que tener en cuenta las siguientes directrices:

1.	Debe tener un buen formato el Excel: La hoja debe tener el siguiente formato:
•	El índice debe estar en la primera línea horizontal
•	No debe haber espacios entre las líneas verticales
•	Evitar que el contenido de las celdas del índice contenga más de 15 caracteres.

3.	Si desea abrir una hoja la cual esté renombrada o no sea la “Hoja1” primero inserte el nombre de la hoja y luego pulse “Buscar Archivo”.
4.	El archivo debe contener al menos una línea más además del índice, para que el programa pueda determinar qué tipo de dato corresponde.
5.	Para el correcto funcionamiento de las funcionalidades “Añadir” “Borrar” y “Modificar” el archivo original debe permanecer cerrado.
6.	
Una vez abierto su archivo de hojas de cálculo:

1.	Buscar: Debe rellenar un campo, los campos correspondientes a booleanos no cuentan y pulsar el botón.
2.	Borrar: Debe rellenar un campo, los campos correspondientes a booleanos no cuentan, y pulsar el botón.
3.	Añadir: Debe rellenar el mayor número de campos y pulsar el botón, los campos referentes a funciones no se pueden rellenar.
4.	Modificar: Primero debe buscar la línea mediante la función buscar o las funciones para obtener la línea anterior o posterior, una vez obtenida la línea modifique el valor del campo que desee y pulse modificar, si no se pulsa el botón se entenderá que no desea modificar este.
5.	Siguiente: Debe haber una línea por delante de la que ya se encuentra, sino este se bloqueara, cuando se realice cualquier otra acción que no sea esta o Anterior se reiniciara la posición volviendo al inicio del archivo.
6.	Anterior: Debe haber una línea por detrás de la que ya se encuentra, sino este se bloqueara, cuando se realice cualquier otra acción que no sea esta o Siguiente se reiniciara la posición volviendo al inicio del archivo.
7.	Salir: Solo si desea cambiar de archivo o si desea recargar el original debido a nuevos campos o líneas insertadas manualmente desde fuera de la aplicación.
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
This application will automatically create a form using a spreadsheet file.
To make good use of the application, please consider the following guidelines:

1. The Excel file must have a proper format:
• The index should be in the first horizontal row.
• There should be no spaces between vertical lines.
• Avoid having index cell content with more than 15 characters.

3. If you want to open a sheet that has been renamed or is not named "Sheet1," first enter the sheet name and then click on "Search File."
4. The file must contain at least one additional row besides the index, so the program can determine the data type.
4.For the proper functioning of the "Add," "Delete," and "Modify" features, the original file must remain closed.

Once your spreadsheet file is opened:

1. Search: Fill in a field, excluding boolean fields, and click the button.
2. Delete: Fill in a field, excluding boolean fields, and click the button.
3. Add: Fill in as many fields as possible and click the button. Function-related fields cannot be filled.
4. Modify: First, search for the line using the search function or functions to retrieve the previous or next line. Once you have obtained the line, modify the value of the desired field and click "Modify." If the button is not pressed, it will be understood that you do not wish to modify it.
5. Next: There must be a line ahead of the current one; otherwise, it will be blocked. When any action other than "Previous" is performed, the position will reset to the beginning of the file.
6. Previous: There must be a line behind the current one; otherwise, it will be blocked. When any action other than "Next" is performed, the position will reset to the beginning of the file.
7. Exit: Use this option only if you want to change the file or reload the original file due to new manually inserted fields or lines from outside the application.
