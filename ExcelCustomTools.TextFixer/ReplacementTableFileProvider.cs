using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelCustomTools.TextFixer
{
    public class ReplacementTableFileProvider : IReplacementTableProvider
    {
        private const string FileName = "ExcelCustomTools_TextFixer.txt";

        public Dictionary<string, string> GetReplacementTable()
        {
            var replacementTable = new Dictionary<string, string>();

            var filePath
                = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), FileName);

            var fileArray = File.ReadAllLines(filePath);

            for (int i = 0; i < fileArray.Length; i++)
            {
                var line = fileArray[i];
                var items = line.Split('\t');

                if (items.Length == 2)
                {
                    replacementTable.Add(items[0], items[1]);
                }
            }

            return replacementTable;
        }
    }
}
