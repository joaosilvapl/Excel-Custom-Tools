using System;
using System.Collections.Generic;

namespace ExcelCustomTools.TextFixer
{
    public class TextFixer : ITextFixer
    {
        private readonly IReplacementTableProvider _replacementTableProvider;
        private Dictionary<string, string> _replacementTable;
        private bool _initialized;

        public TextFixer(IReplacementTableProvider replacementTableProvider)
        {
            this._replacementTableProvider = replacementTableProvider ?? throw new ArgumentNullException(nameof(replacementTableProvider));
        }

        public void Initialize()
        {
            if (!this._initialized)
            {
                this._replacementTable = this._replacementTableProvider.GetReplacementTable();

                this._initialized = true;
            }
        }

        public FixResult Fix(string text)
        {
            if (!this._initialized)
            {
                throw new ApplicationException("Not initialized");
            }

            var newText = text;

            if (!String.IsNullOrWhiteSpace(newText))
            {
                //TODO: Optimize string replacement and comparison

                foreach (KeyValuePair<string, string> entry in _replacementTable)
                {
                    newText = newText.Replace(entry.Key, entry.Value);
                }
            }

            var textChanged = String.Compare(text, newText, StringComparison.Ordinal) != 0;

            return new FixResult
            {
                NewText = newText,
                TextChanged = textChanged
            };
        }

        public class FixResult
        {
            public string NewText;
            public bool TextChanged;
        }
    }
}
