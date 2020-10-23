using System.Collections.Generic;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class RangeData : IDataWriter
    {
        public IEnumerable<IRowData> Rows { get; set; }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            var rowStart = rowIndex;
            foreach (var row in Rows)
            {
                row.Write(writer, rowIndex, colIndex);
                rowIndex++;
            }
        }
    }
}
