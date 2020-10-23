using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class DataCoordWriter : IDataCoordWriter
    {
        private readonly IDataWriter dataWriter;

        public DataCoordWriter(IDataWriter dataWriter)
        {
            this.dataWriter = dataWriter;
        }

        public int RowIndex { get; set; }

        public int ColIndex { get; set; }

        public string SheetName { get; set; }

        public void Write(ISpreadSheetHandler writer)
        {
            dataWriter.Write(writer, RowIndex, ColIndex);
        }
    }
}
