namespace ExcelCustomTools.TextFixer
{
    public interface ITextFixer
    {
        void Initialize();
        TextFixer.FixResult Fix(string text);
    }
}