using Tools.SpreadSheets.Factories;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Light
{
    public class SpreadSheetLightHandlerFactory : ISpreadSheetHandlerFactory
    {
        public static ISpreadSheetGenerator CreateGenerator()
        {
            return new SpreadSheetGenerator(new SpreadSheetLightHandlerFactory(), new SpreadSheetWriterFactory());
        }

        public ISpreadSheetHandler CreateHandler(string path)
        {
            return new SpreadSheetLiteHandler(new SpreadsheetLight.SLDocument(path));
        }
    }
}
