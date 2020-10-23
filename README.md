# SpreadSheetLightDataset

## Uso del Alias o Nombres de rango 
Utiliza los Alias o Nombres de celda para relacionar un set de datos con la posición en la que debe incluirse:



 

 

### Rango con columnas fijas
Recibe un set de datos y los impacta en un Excel existente agregando filas de forma dinamica según la cantidad de filas insertadas.


### Rango con columnas variables
Recibe un set de datos y los impacta en un Excel existente agregando filas y columnas de forma dinamica según la cantidad de filas y columnas insertadas.


 

## Desplazamiento de datos
Existen casos en los que vamos a necesitar dezplazar el contenido de una fila a la derecha, en lugar de insertar nuevas columnas. Ya que, al insertar nuevas colunmas, estamos afectando a toda la hoja y puede que estemos afectando a otras celdas, por encima o debajo, de la que queremos editar.

 

Supongamos el caso en el que tenemos que insertar el detalle de las compras, en el siguiente template:


Ya se completó la tabla de clientes, ahora debería completarse la tabla de compras


Cuando el proceso intente completar la compras y agregue nuevas columnas, estará afectando la tabla de clientes.

 

 

Para solucionar este problema, en lugar de insertar columnas, lo que debemos hacer es desplazar las celdas hacia la derecha, a medida que vamos agregando nueva información.


 


 

 

Más información
SpreadSheet - Generar datos y comportamientos de escritura
 

## Diagrama técnico
Tool.SpreadSheet
