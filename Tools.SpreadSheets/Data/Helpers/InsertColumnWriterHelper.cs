using System.Collections.Generic;
using System.Linq;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class InsertColumnWriterHelper
    {
        public static DynamicRowData ToDynamicRowData(this IEnumerable<IDataWriter> cells)
        {
            return new DynamicRowData
            {
                Cells = cells
            };
        }

    }
}
