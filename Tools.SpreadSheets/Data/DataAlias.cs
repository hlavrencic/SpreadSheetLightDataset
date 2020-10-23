using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class DataAlias : IDataAliasWriter
    {
        private readonly IDataWriter dataWriter;

        public DataAlias(IDataWriter dataWriter)
        {
            this.dataWriter = dataWriter;
        }

        public string Alias { get; set; }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            dataWriter.Write(writer, rowIndex, colIndex);
        }
    }
}