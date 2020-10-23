using System.Collections.Generic;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class DynamicRowWriterHelper
    {
        public static RangeData ToDynamicRangeData(this IEnumerable<IEnumerable<IDataWriter>> rowsData)
        {
            return rowsData.ToRowData().ToAlgo().ToRangeData();
        }

        public static RowCopyData ToRowCopyData(this IEnumerable<IDataWriter> cells)
        {
            return new RowCopyData(cells.ToRowData());
        }

        public static RangeData ToDynamicRangeData(this IEnumerable<IRowData> rowsData)
        {
            return rowsData.ToRangeData();
        }

        private static IEnumerable<IRowData> ToAlgo(this IEnumerable<IRowData> rowsData)
        {
            IRowData ultimo = null;
            var count = 0;
            foreach (var row in rowsData)
            {
                if (count == 0)
                {
                    yield return row;
                }
                else
                {
                    if (ultimo == null)
                    {
                        ultimo = row;
                    }
                    else if (ultimo != row)
                    {
                        yield return new RowCopyData(ultimo);
                        ultimo = row;
                    }
                }

                count++;
            }

            yield return ultimo;
        }
    }
}
