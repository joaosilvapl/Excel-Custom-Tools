using System.Collections.Generic;

namespace ExcelCustomTools.TextFixer
{
    public interface IReplacementTableProvider
    {
        Dictionary<string, string> GetReplacementTable();
    }
}