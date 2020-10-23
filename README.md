# SpreadSheetLightDataset

## Uso del Alias o Nombres de rango 
Utiliza los Alias o Nombres de celda para relacionar un set de datos con la posición en la que debe incluirse:


|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/89a79e37-af02-42ad-802b-e55ea5db8bb0.png)|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/58b803fd-b996-4e42-b083-8f51e4dc9c74.png)|
|--|--|
 

 

### Rango con columnas fijas
Recibe un set de datos y los impacta en un Excel existente agregando filas de forma dinamica según la cantidad de filas insertadas.

![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/6d543c89-a04f-4fd6-a274-eb7455a856cb.png)


### Rango con columnas variables
Recibe un set de datos y los impacta en un Excel existente agregando filas y columnas de forma dinamica según la cantidad de filas y columnas insertadas.

![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/6cfe289f-7f10-468a-a685-4e940b44667f.png)
  
 
 
## Desplazamiento de datos
Existen casos en los que vamos a necesitar dezplazar el contenido de una fila a la derecha, en lugar de insertar nuevas columnas. Ya que, al insertar nuevas colunmas, estamos afectando a toda la hoja y puede que estemos afectando a otras celdas, por encima o debajo, de la que queremos editar.
 
 
 
Supongamos el caso en el que tenemos que insertar el detalle de las compras, en el siguiente template:
|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/7c670c16-fcb5-4931-94de-e7d107da387b.png)|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/4baded35-2ff3-4d87-8398-fb019bf9dd33.png)|
|--|--|
|Ya se completó la tabla de clientes, ahora debería completarse la tabla de compras|Cuando el proceso intente completar la compras y agregue nuevas columnas, estará afectando la tabla de clientes.|
  
  
  
Para solucionar este problema, en lugar de insertar columnas, lo que debemos hacer es desplazar las celdas hacia la derecha, a medida que vamos agregando nueva información.
|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/648b6a5b-92fd-41db-aeb4-de98c9dd1058.png)|![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/34c80d84-aa10-48b4-8d60-bd0e338fc3e8.png)|
|--|--|

## Diagrama técnico

![](https://raw.githubusercontent.com/hlavrencic/SpreadSheetLightDataset/master/doc/img/diagrama1.png) 
 
 
  
   
   
   
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
