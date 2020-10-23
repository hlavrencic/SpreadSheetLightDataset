using System;
using System.Collections.Generic;
using System.Linq;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class ShiftRowData : IRowData
    {
        private readonly IRowData rowData;

        public ShiftRowData(IRowData rowData)
        {
            this.rowData = rowData;
        }

        public IEnumerable<IDataWriter> Cells => GetShiftedCells();

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

        private IEnumerable<IDataWriter> GetShiftedCells()
        {
            var firstSent = false;
            foreach (var cell in rowData.Cells)
            {
                if (firstSent)
                {
                    yield return new ShiftCellData(cell);
                }
                else
                {
                    yield return cell;
                    firstSent = true;
                }
            }
        }
    }
}
