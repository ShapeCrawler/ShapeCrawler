using FluentAssertions;
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
        slideMaster.SlideNumber!.Font.Size = 30;

        // Assert
        pres.Save();
        pres = SCPresentation.Open(new MemoryStream(pres.BinaryData));
        slideMaster = pres.SlideMasters[0];
        slideMaster.SlideNumber!.Font.Size.Should().Be(30);
    }
    
    [Test]
    public void WIP()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slideMaster = pres.SlideMasters[0];

        // Add slide number
        pres.HeaderAndFooter.AddSlideNumber();
  
        // Change slide number color
        var green = SCColor.FromHex("00FF00");
        slideMaster.SlideNumber!.Font.Color = green;
  
        // Change slide number size
        slideMaster.SlideNumber!.Font.Size = 30;
        
        // Change slide number location
        slideMaster.SlideNumber.X -= 400;

        SaveResult(pres);
    }
}