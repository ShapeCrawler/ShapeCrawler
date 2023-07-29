using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class SlideMasterTests : SCTest
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
    
    [Test]
    public void SlideNumber_Font_Size_Setter()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slideMaster = pres.SlideMasters[0];

        // Act
        pres.HeaderAndFooter.AddSlideNumber();
        // slideMaster.Sli      deNumber!.Font.Size = 30;
        SaveResult(pres);

        // Assert
        // pres.Save();
        // pres = SCPresentation.Open(new MemoryStream(pres.BinaryData));
        // slideMaster = pres.SlideMasters[0];
        // Assert.That(slideMaster.SlideNumber!.Font.Size, Is.EqualTo(30));
    }
}