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


# Generar datos y comportamientos de escritura

## Objeto Celda
Todo dato que desee volcarse a la planilla debe ser convertido en Celda. Para eso debe utilizarse el helper .ToCell()

Ejemplo:
```csharp
DateTime fecha;
IDataWriter celda = fecha.ToCell();
```

Cualquier tipo de dato puede ser convertido en celda:
```csharp
double precio;
IDataWriter celda = precio.ToCell();
```
Cada celda conservará su tipo de origen e impactará el valor como corresponda.

## Objeto Fila
Una lista de celdas puede agruparse en una fila. Que luego se escribira como tal en la planilla.

Ejemplo:
```csharp
IEnumerable<IDataWriter> celdas;
IRowData fila = celdas.ToRowData();
```
 
## Objeto Rango

Una lista de filas puede agruparse en un rango:
```csharp
IEnumerable<IRowData> filas;
RangeData rango = filas.ToRangeData();
```
 
También es posible crear un rango a partir de una lista de listas de celdas:
```csharp
IEnumerable<IEnumerable<IDataWriter>> celdas;
RangeData rango = celdas.ToRangeData();
```

## Filas y rangos con comportamiento
Para insertar nuevas filas antes de escribir la información:

```csharp
IEnumerable<IEnumerable<IDataWriter>> celdas;
RangeData rango = celdas.ToDynamicRangeData();
```

Para crear una fila de celdas que va insertando nuevas columnas a medida que escribe la información:
```csharp
IEnumerable<IDataWriter> celdas;
DynamicRowData rango = celdas.ToDynamicRowData();
```

Para crear una fila o rango de celdas que vaya desplazando celdas a la derecha a medida que inserta nueva información:
```csharp
IEnumerable<IDataWriter> celdas;
ShiftRowData rango = celdas.ToShiftRowData();
```

## Para rangos:

```csharp
IEnumerable<IEnumerable<IDataWriter>> filas;
RangeData rango = filas.ToShiftRangeData();
```

## Asociar alias
Para que una celda, fila o rango pueda escribirse en un excel, debe asociarse a un alias:

```csharp
// Celda
DateTime fecha;
IDataAliasWriter celda = fecha.ToCell("alias1");

//Fila
IDataAliasWriter fila = celdas.ToRowData("alias2"); 

//Rango
IDataAliasWriter rango = filas.ToRangeData("alias3");
```

Todos los metodos estaticos tienen una sobrecarga que permite ingresar el alias.
