using Tools.SpreadSheets.Handlers;
using Tools.SpreadSheets.Writers;

namespace Tools.SpreadSheets.Factories
{
    public class SpreadSheetWriterFactory : ISpreadSheetWriterFactory
    {
        public ISpreadSheetWriter Create(ISpreadSheetHandler handler)
        {
            return new SpreadSheetWriter(handler);
        }
    }
}
