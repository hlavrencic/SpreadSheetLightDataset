using System.Collections.Generic;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class DynamicRowData : IDataWriter, IRowData
    {
        public IEnumerable<IDataWriter> Cells { get; set; }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            var col = colIndex;
            foreach (var cell in Cells)
            {
                if(col != colIndex)
                {
                    writer.CopyCol(col - 1);
                }

                // Escribe la celda por celda
                cell.Write(writer, rowIndex, col);

                // Se desplaza por cada columna
                col++;
            }
        }
    }
}
