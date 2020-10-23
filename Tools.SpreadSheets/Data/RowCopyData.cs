using System.Collections.Generic;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class RowCopyData : IRowData
    {
        private readonly IRowData rowData;

        public RowCopyData(IRowData rowData)
        {
            this.rowData = rowData;
        }

        public IEnumerable<IDataWriter> Cells => rowData.Cells;

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            writer.CopyRow(rowIndex - 1);
            rowData.Write(writer, rowIndex, colIndex);
        }
    }
}
