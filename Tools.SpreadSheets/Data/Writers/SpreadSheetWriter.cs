using System;
using System.Collections.Generic;
using System.Linq;
using Tools.SpreadSheets.Data;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Writers
{
    public class SpreadSheetWriter : ISpreadSheetWriter
    {
        private readonly ISpreadSheetHandler spreadSheetHandler;

        public SpreadSheetWriter(ISpreadSheetHandler spreadSheetHandler)
        {
            this.spreadSheetHandler = spreadSheetHandler;
        }

        public void Write(IEnumerable<IDataAliasWriter> dataAliasWriters)
        {
            var filteredAliases = dataAliasWriters
                .Select(GetCoords)
                .Where(data => data != null)
                .OrderBy(s => s.SheetName)
                .OrderByDescending(c => c.RowIndex );
            string sheetName = null;
            foreach (var item in filteredAliases)
            {
                // Si la hoja actual es disitinta, me posiciono nuevamente.
                if (string.IsNullOrEmpty(sheetName) || sheetName != item.SheetName)
                {
                    sheetName = item.SheetName;
                    spreadSheetHandler.SelectWorksheet(item.SheetName);
                }

                item.Write(spreadSheetHandler);
            }
        }

        private IDataCoordWriter GetCoords(IDataAliasWriter dataAliasWriters)
        {
            var pos = spreadSheetHandler.GetPosition(dataAliasWriters);

            if(pos == null)
            {
                return null;
            }

            return new DataCoordWriter(dataAliasWriters)
            {
                ColIndex = pos.ColIndex,
                RowIndex = pos.RowIndex,
                SheetName = pos.SheetName
            };
        }
    }
}
