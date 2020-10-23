using System.Collections.Generic;
using System.Linq;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class ShiftWriterHelper
    {
        public static RangeData ToShiftRangeData(this IEnumerable<IEnumerable<IDataWriter>> dataWriters)
        {
            return dataWriters.ToShiftRowData().ToRangeData();
        }

        public static IEnumerable<ShiftRowData> ToShiftRowData(this IEnumerable<IEnumerable<IDataWriter>> dataWriters)
        {
            return dataWriters.Select(dataWriter => dataWriter.ToShiftRowData());
        }

        public static IEnumerable<ShiftRowData> ToShiftRowData(this IEnumerable<IRowData> rowsData)
        {
            return rowsData.Select(rowData => rowData.ToShiftRowData());
        }

        public static ShiftRowData ToShiftRowData(this IEnumerable<IDataWriter> cells)
        {
            return cells.ToRowData().ToShiftRowData();
        }

        public static ShiftRowData ToShiftRowData(this IRowData rowData)
        {
            return new ShiftRowData(rowData);
        }
    }
}
