namespace Tools.SpreadSheets.Data
{
    /// <summary>
    /// Objeto que puede escribirse y tiene alias
    /// </summary>
    public interface IDataAliasWriter : IDataWriter
    {
        string Alias { get; }
    }
}
