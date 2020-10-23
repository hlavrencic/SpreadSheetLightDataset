using System;
using Tools.SpreadSheets.Data;

namespace Tools.SpreadSheets.Handlers
{
    /// <summary>
    /// Conecta con la funcionalidad de manejo de Excel 
    /// de alguna libreria externa y se encarga de abstraer 
    /// todo de la misma.
    /// </summary>
    public interface ISpreadSheetHandler : IDisposable
    {
        /// <summary>
        /// Escribe el dato en la coordenada especificada.
        /// Conversion de tipo automatica.
        /// </summary>
        void Write<TData>(TData data, int rowIndex, int colIndex);

        /// <summary>
        /// Crea una copia de la fila y la inserta debajo.
        /// Conserva formatos y formulas. Adaptandolas a 
        /// la nueva fila.
        /// </summary>
        /// <param name="rowSourceIndex">Nro de fila que comienza en 1</param>
        void CopyRow(int rowSourceIndex);

        /// <summary>
        /// Crea una copia de la columna y la inserta a su derecha.
        /// Conserva formatos y formulas. Adaptandolas a la nueva columna.
        /// </summary>
        /// <param name="colSourceIndex">Nro de columna que comienza en 1</param>
        void CopyCol(int colSourceIndex);

        /// <summary>
        /// Desplaza todas las celdas hacia la derecha
        /// a partir de cierta columna indicada
        /// </summary>
        void ShiftRight(int rowIndex, int colIndex);

        /// <summary>
        /// Devuelve un nuevo dato con las coordenadas del alias
        /// </summary>
        /// <param name="data">Celda, Fila o Rango con Alias</param>
        /// <returns></returns>
        IDataCoordWriter GetPosition(IDataAliasWriter data);

        /// <summary>
        /// Selecciona la hoja activa.
        /// Donde se realizan las acciones de escritura.
        /// </summary>
        /// <param name="workSheetName"></param>
        /// <returns></returns>
        bool SelectWorksheet(string workSheetName);

        /// <summary>
        /// Guarda el archivo
        /// </summary>
        /// <param name="path"></param>
        void Save(string path);

        void RenameWorkSheet(string name);
    }
}
