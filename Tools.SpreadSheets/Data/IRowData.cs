using System.Collections.Generic;

namespace Tools.SpreadSheets.Data
{
    public interface IRowData : IDataWriter
    {
        IEnumerable<IDataWriter> Cells { get; }
    }
}