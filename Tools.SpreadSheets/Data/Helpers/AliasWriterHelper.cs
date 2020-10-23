using System.Collections.Generic;

namespace Tools.SpreadSheets.Data.Helpers
{
    public static class AliasWriterHelper
    {
        public static IDataAliasWriter ToShiftRangeData(this IEnumerable<IEnumerable<IDataWriter>> dataWriters, string alias)
        {
            return dataWriters.ToShiftRangeData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToShiftRowData(this IEnumerable<IDataWriter> dataWriters, string alias)
        {
            return dataWriters.ToShiftRowData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToDynamicRangeData(this IEnumerable<IEnumerable<IDataWriter>> rowsData, string alias)
        {
            return rowsData.ToDynamicRangeData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToDynamicRangeData(this IEnumerable<IRowData> rowsData, string alias)
        {
            return rowsData.ToDynamicRangeData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToRangeData(this IEnumerable<IRowData> rowsData, string alias)
        {
            return rowsData.ToRangeData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToRangeData(this IEnumerable<IEnumerable<IDataWriter>> rowsData, string alias)
        {
            return rowsData.ToRangeData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToDynamicRowData(this IEnumerable<IDataWriter> cells, string alias)
        {
            return cells.ToDynamicRowData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToRowData(this IEnumerable<IDataWriter> cells, string alias)
        {
            return cells.ToRowData().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToCell<TData>(this TData data, string alias)
        {
            return data.ToCell().ToAliasWriter(alias);
        }

        public static IDataAliasWriter ToAliasWriter(this IDataWriter writer, string alias)
        {
            return new DataAlias(writer)
            {
                Alias = alias
            };
        }
    }
}
