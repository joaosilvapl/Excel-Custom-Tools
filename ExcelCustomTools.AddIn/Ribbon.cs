using System;
using System.Windows.Forms;
using ExcelCustomTools.TextFixer;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelCustomTools.AddIn
{
    public partial class Ribbon
    {
        private const string TextFixerMessageBoxCaption = "Text fixer";

        private readonly DependencyContainer _dependencyContainer = new DependencyContainer();
        private ITextFixer _textFixer;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            this._dependencyContainer.Initialize();

            this._textFixer = this._dependencyContainer.Resolve<ITextFixer>();
            this._textFixer.Initialize();
        }

        private void btnFixText_Click(object sender, RibbonControlEventArgs e)
        {
            var currentSelection = Globals.ThisAddIn.GetSelection();

            const int maxIterations = 100000;
            var index = 0;
            var fixCount = 0;

            foreach (Range cell in currentSelection.Cells)
            {
                try
                {
                    var currentText = cell.Value2?.ToString();

                    var fixResult = this._textFixer.Fix(currentText);

                    if (fixResult.TextChanged)
                    {
                        cell.Value2 = fixResult.NewText;
                        fixCount++;
                    }

                    if (index++ > maxIterations)
                    {
                        MessageBox.Show($@"Operation aborted after {maxIterations} iterations", TextFixerMessageBoxCaption);
                        break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($@"An exception has occurred: {ex}", TextFixerMessageBoxCaption);
                }
            }

            MessageBox.Show($@"Cells process: {index}{Environment.NewLine}Fixed cells: {fixCount}", TextFixerMessageBoxCaption);
        }
    }
}
