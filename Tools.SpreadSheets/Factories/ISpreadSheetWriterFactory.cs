using Tools.SpreadSheets.Handlers;
using Tools.SpreadSheets.Writers;

namespace Tools.SpreadSheets.Factories
{
    public interface ISpreadSheetWriterFactory
    {
        ISpreadSheetWriter Create(ISpreadSheetHandler handler);
    }
}