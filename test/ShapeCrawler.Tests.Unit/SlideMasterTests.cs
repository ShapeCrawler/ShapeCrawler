using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class SlideMasterTests
{
    [Test]
    [PresentationData("new")]
    [PresentationData("023.pptx")]
    public void SlideNumber_Font_Color_Setter(IPresentation pres)
    {
        // Arrange
        var slideMaster = pres.SlideMasters[0];
        var green = SCColor.FromHex("00FF00");

        // Act
        slideMaster.SlideNumber!.Font.Color = green;

        // Assert
        Assert.That(slideMaster.SlideNumber.Font.Color.Hex, Is.EqualTo("00FF00"));
    }
}