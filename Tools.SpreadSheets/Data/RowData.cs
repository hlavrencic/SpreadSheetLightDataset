using System.Collections.Generic;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class RowData : IDataWriter, IRowData
    {
        public IEnumerable<IDataWriter> Cells { get; set; }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            var col = colIndex;
            foreach (var cell in Cells)
            {
                // Escribe la celda por celda
                cell.Write(writer, rowIndex, col);

                // Se desplaza por cada columna
                col++;
            }
        }
    }
}
