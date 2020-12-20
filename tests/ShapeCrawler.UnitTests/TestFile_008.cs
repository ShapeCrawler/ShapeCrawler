using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Models;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Placeholders;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_008
    {
        [Fact]
        public void ShapeTextBody_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._008);

            // ACT
            var shapes = pre.Slides.Single().Shapes.OfType<ShapeEx>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
           
            pre.Close();

            // ASSERT
            Assert.False(sh36.HasTextFrame);
            Assert.True(sh37.HasTextFrame);
            Assert.Equal("P1t1 P1t2", sh37.TextFrame.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextFrame.Paragraphs[1].Text);
        }

        [Fact]
        public void Placeholder_DateTime_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextFrame = sp3.HasTextFrame;
            var text = sp3.TextFrame.Text;
            var phType = sp3.PlaceholderType;
            var x = sp3.X;

            pre.Close();

            // ASSERT
            Assert.True(hasTextFrame);
            Assert.Equal("25.01.2020", text);
            Assert.Equal(PlaceholderType.DateAndTime, phType);
            Assert.Equal(628650, x);
        }

        [Fact]
        public void Get_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);
            var sldLtPart = sldPart.SlideLayoutPart;
            var phService = new PlaceholderService(sldLtPart);

            // ACT
            var type = phService.TryGetLocation(spId3).PlaceholderType;

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, type);
        }

        [Fact]
        public void GetPlaceholderType_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);
            var placeholderService = new PlaceholderService(sldPart.SlideLayoutPart);

            // ACT
            var phXml = placeholderService.CreatePlaceholderData(spId3);

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, phXml.PlaceholderType);
        }
    }
}
