using System.Collections.Generic;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class RowDataHelper
    {
        public static RangeData ToRangeData(this IEnumerable<IEnumerable<IDataWriter>> cells)
        {
            return cells.ToRowData().ToRangeData();
        }

        public static RangeData ToRangeData(this IEnumerable<IRowData> rowsData)
        {
            return new RangeData
            {
                Rows = rowsData
            };
        }
    }
}
