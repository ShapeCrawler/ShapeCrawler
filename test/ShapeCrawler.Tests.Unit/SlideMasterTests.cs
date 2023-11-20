using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class SlideMasterTests : SCTest
{
    [Test]
    [Presentation("new")]
    [Presentation("023.pptx")]
    public void SlideNumber_Font_Color_Setter(IPresentation pres)
    {
        // Arrange
        var slideMaster = pres.SlideMasters[0];
        var green = Color.FromHex("00FF00");

        // Act
        slideMaster.SlideNumber!.Font.Color = green;

        // Assert
        Assert.That(slideMaster.SlideNumber.Font.Color.Hex, Is.EqualTo("00FF00"));
    }
    
    [Test]
    public void SlideNumber_Font_Size_Setter()
    {
        // Arrange
        var pres = new Presentation();
        var slideMaster = pres.SlideMasters[0];

        // Act
        pres.Footer.AddSlideNumber();
        slideMaster.SlideNumber!.Font.Size = 30;

        // Assert
        pres.Save();
        pres = new Presentation(new MemoryStream(pres.AsByteArray()));
        slideMaster = pres.SlideMasters[0];
        slideMaster.SlideNumber!.Font.Size.Should().Be(30);
    }
}