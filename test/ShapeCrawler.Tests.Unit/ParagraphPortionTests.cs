using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class ParagraphPortionTests : SCTest
{
    [Test]
    public void Text_Getter_returns_text_of_paragraph_portion()
    {
        // Arrange
        var pptx = StreamOf("009_table");
        var pres = new Presentation(pptx);
        IParagraphPortion portion = ((ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
            .TextFrame
            .Paragraphs[0].Portions[0];

        // Act
        string paragraphPortionText = portion.Text;

        // Assert
        paragraphPortionText.Should().BeEquivalentTo("0:0_p1_lvl1");
    }
}