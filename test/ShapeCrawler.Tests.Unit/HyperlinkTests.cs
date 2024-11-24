// using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class HyperlinkTests
{
    // [Fact]
    // public void Should_Create_Inner_Slide_Hyperlink()
    // {
    //     // Arrange
    //     var pre = new Presentation(@"c:\temp\link.pptx");
    //     var slide = pre presentation.Slides[0];
    //     var shape = slide.AddShape(ShapeType.Rectangle, 100, 100, 100, 100);
    //     var text = shape.TextFrame.Text;
    //     text.Value = "Go to Slide 2";
    //
    //     // Act
    //     text.Hyperlink = "slide://2";
    //
    //     // Assert
    //     Assert.Equal("slide://2", text.Hyperlink);
    // }

    // [Fact]
    // public void Should_Throw_Exception_For_Invalid_Slide_Number()
    // {
    //     // Arrange
    //     using var presentation = TestPresentation.Create();
    //     var slide = presentation.Slides[0];
    //     var shape = slide.AddShape(ShapeType.Rectangle, 100, 100, 100, 100);
    //     var text = shape.TextFrame.Text;
    //     text.Value = "Invalid Link";
    //
    //     // Act & Assert
    //     var ex = Assert.Throws<SCException>(() => text.Hyperlink = "slide://999");
    //     Assert.Equal("Invalid slide number: 999", ex.Message);
    // }
}
