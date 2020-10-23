using System.Collections.Generic;
using System.IO;
using Tools.SpreadSheets.Data;
using Tools.SpreadSheets.Factories;

namespace Tools.SpreadSheets
{
    public class SpreadSheetGenerator : ISpreadSheetGenerator
    {
        private readonly ISpreadSheetHandlerFactory spreadSheetHandlerFactory;
        private readonly ISpreadSheetWriterFactory spreadSheetRangeWriterFactory;

        public SpreadSheetGenerator(
            ISpreadSheetHandlerFactory spreadSheetHandlerFactory,
            ISpreadSheetWriterFactory spreadSheetRangeWriterFactory)
        {
            this.spreadSheetHandlerFactory = spreadSheetHandlerFactory;
            this.spreadSheetRangeWriterFactory = spreadSheetRangeWriterFactory;
        }

        public void CreateFromTemplate(string pathTemplate, string newFilePath, params IDataAliasWriter[] dataAliasWriters)
        {
            using (var spreadSheetLiteHandler = spreadSheetHandlerFactory.CreateHandler(pathTemplate))
            {
                var spreadSheetRangeWriter = spreadSheetRangeWriterFactory.Create(spreadSheetLiteHandler);
                spreadSheetRangeWriter.Write(dataAliasWriters);

                spreadSheetLiteHandler.Save(newFilePath);
            }
        }

        public void CreateFromTemplate(string pathTemplate, string newFilePath, string sheetName, params IDataAliasWriter[] dataAliasWriters)
        {
            using (var spreadSheetLiteHandler = spreadSheetHandlerFactory.CreateHandler(pathTemplate))
            {
                var spreadSheetRangeWriter = spreadSheetRangeWriterFactory.Create(spreadSheetLiteHandler);

                spreadSheetLiteHandler.RenameWorkSheet(sheetName);
                spreadSheetRangeWriter.Write(dataAliasWriters);


                spreadSheetLiteHandler.Save(newFilePath);
            }
        }
    }
}
