using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class CellData<TData> : ICellWriter
    {
        public TData Value { get; set; }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            writer.Write(Value, rowIndex, colIndex);
        }
    }
}
