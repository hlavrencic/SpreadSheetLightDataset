using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    /// <summary>
    /// Objeto con las coordenadas de escritura y puede escribirse.
    /// </summary>
    public interface IDataCoordWriter
    {
        int RowIndex { get; }

        int ColIndex { get; }

        string SheetName { get; }

        void Write(ISpreadSheetHandler writer);
    }
}
