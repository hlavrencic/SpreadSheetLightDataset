using System.Collections.Generic;
using Tools.SpreadSheets.Data;

namespace Tools.SpreadSheets.Writers
{
    public interface ISpreadSheetWriter
    {
        void Write(IEnumerable<IDataAliasWriter> dataAliasWriters);
    }
}