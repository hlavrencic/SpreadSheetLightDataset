using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    /// <summary>
    /// Objeto que puede escribirse
    /// </summary>
    public interface IDataWriter
    {
        void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex);
    }
}
