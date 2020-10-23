using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Factories
{
    public interface ISpreadSheetHandlerFactory
    {
        ISpreadSheetHandler CreateHandler(string path);
    }
}
