using System.Collections.Generic;
using System.Linq;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class CellDataHelper
    {
        

        public static IEnumerable<RowData> ToRowData(this IEnumerable<IEnumerable<IDataWriter>> cells)
        {
            return cells.Select(cell => cell.ToRowData());
        }

        public static RowData ToRowData(this IEnumerable<IDataWriter> cells)
        {
            return new RowData
            {
                Cells = cells
            };
        }

        public static CellData<TData> ToCell<TData>(this TData data)
        {
            return new CellData<TData>
            {
                Value = data
            };
        }
    }
}
