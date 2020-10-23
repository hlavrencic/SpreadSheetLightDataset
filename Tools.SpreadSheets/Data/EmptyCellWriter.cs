using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class EmptyCellWriter : ICellWriter
    {
        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            // Do nothing.
        }
    }
}
