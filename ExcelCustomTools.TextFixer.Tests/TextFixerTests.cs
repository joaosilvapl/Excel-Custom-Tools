using System;
using System.Collections.Generic;
using NUnit.Framework;
using Rhino.Mocks;

namespace ExcelCustomTools.TextFixer.Tests
{
    [TestFixture]
    public class TextFixerTests
    {
        [Test]
        public void Constructor_if_null_replacementTableProvider_should_throw_ArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => this.GetTextFixer(null));
        }

        [Test]
        public void Fix_if_not_initialized_ApplicationException_should_be_thrown()
        {
            var textFixer = this.GetTextFixer(MockRepository.GenerateStub<IReplacementTableProvider>());

            Assert.Throws<ApplicationException>(() => textFixer.Fix(null));
        }

        [Test]
        public void Fix_if_no_replacement_should_return_TextChanged_false()
        {
            //Arrange
            const string inputText = "xyz";

            Dictionary<string, string> replacementTable = new Dictionary<string, string>
            {
                {"aaa", "bbb"}
            };

            var fakeReplacementTableProvider = MockRepository.GenerateStub<IReplacementTableProvider>();
            fakeReplacementTableProvider.Stub(x => x.GetReplacementTable()).Return(replacementTable);

            var textFixer = this.GetTextFixer(fakeReplacementTableProvider);
            textFixer.Initialize();

            //Act
            var result = textFixer.Fix(inputText);

            //Assert
            Assert.That(result.TextChanged, Is.False);
        }

        [Test]
        public void Fix_if_replacement_should_return_TextChanged_true_and_correct_NewText()
        {
            //Arrange
            const string inputText = "tyxaaazzz";
            const string expectedNewText = "tyxbbbzzz";

            Dictionary<string, string> replacementTable = new Dictionary<string, string>
            {
                {"aaa", "bbb"}
            };

            var fakeReplacementTableProvider = MockRepository.GenerateStub<IReplacementTableProvider>();
            fakeReplacementTableProvider.Stub(x => x.GetReplacementTable()).Return(replacementTable);

            var textFixer = this.GetTextFixer(fakeReplacementTableProvider);
            textFixer.Initialize();

            //Act
            var result = textFixer.Fix(inputText);

            //Assert
            Assert.That(result.TextChanged, Is.True);
            Assert.That(result.NewText, Is.EqualTo(expectedNewText));
        }

        [Test]
        public void Fix_if_multiple_replacements_should_return_TextChanged_true_and_correct_NewText()
        {
            //Arrange
            const string inputText = "tyxaaazzzDEF765";
            const string expectedNewText = "tyxbbbzzzopq765";

            Dictionary<string, string> replacementTable = new Dictionary<string, string>
            {
                {"aaa", "bbb"},
                {"DEF", "opq"}
            };

            var fakeReplacementTableProvider = MockRepository.GenerateStub<IReplacementTableProvider>();
            fakeReplacementTableProvider.Stub(x => x.GetReplacementTable()).Return(replacementTable);

            var textFixer = this.GetTextFixer(fakeReplacementTableProvider);
            textFixer.Initialize();

            //Act
            var result = textFixer.Fix(inputText);

            //Assert
            Assert.That(result.TextChanged, Is.True);
            Assert.That(result.NewText, Is.EqualTo(expectedNewText));
        }

        public TextFixer GetTextFixer(IReplacementTableProvider replacementTableProvider)
        {
            return new TextFixer(replacementTableProvider);
        }
    }
}
