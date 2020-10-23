using System.Collections.Generic;
using Tools.SpreadSheets.Handlers;

namespace Tools.SpreadSheets.Data
{
    public class ShiftCellData : IDataWriter
    {
        private readonly IDataWriter dataWriter;

        public ShiftCellData(IDataWriter dataWriter)
        {
            this.dataWriter = dataWriter;
        }

        public void Write(ISpreadSheetHandler writer, int rowIndex, int colIndex)
        {
            writer.ShiftRight(rowIndex, colIndex);
            dataWriter.Write(writer, rowIndex, colIndex);
        }
    }
}
