using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NSubstitute;
using SlideXML.Enums;
using SlideXML.Extensions;
using SlideXML.Models.Elements;
using SlideXML.Models.Settings;
using SlideXML.Services;
using SlideXML.Services.Placeholders;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains unit tests for the <see cref="ElementFactory"/> class.
    /// </summary>
    public class ElementFactoryTests
    {
        [Fact]
        public void CreateShape_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubXmlShape = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>().Single(s => s.GetId() == 36);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubXmlShape,
                ElementType = ShapeType.AutoShape
            };
            var mockPhService = Substitute.For<IPlaceholderService>();
            var creator = new ElementFactory(sldPart);
            var stubPhDic = new Dictionary<int, PlaceholderSL>();
            var mockPreSetting = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSetting);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(ShapeType.AutoShape, element.Type);
            Assert.Equal(3291840, element.X);
            Assert.Equal(274320, element.Y);
            Assert.Equal(1143000, element.Width);
            Assert.Equal(1143000, element.Height);
        }

        [Fact]
        public void CreatePicture_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubXmlPic = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.Picture>().Single();
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubXmlPic,
                ElementType = ShapeType.Picture
            };
            var mockPhService = Substitute.For<IPlaceholderService>();
            var creator = new ElementFactory(sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(ShapeType.Picture, element.Type);
            Assert.Equal(4663440, element.X);
            Assert.Equal(1005840, element.Y);
            Assert.Equal(2315880, element.Width);
            Assert.Equal(2315880, element.Height);
        }

        [Fact]
        public void CreateTable_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(e => e.GetId() == 38);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubGrFrame,
                ElementType = ShapeType.Table
            };
            var mockPhService = Substitute.For<IPlaceholderService>();
            var creator = new ElementFactory(sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(ShapeType.Table, element.Type);
            Assert.Equal(453240, element.X);
            Assert.Equal(3417120, element.Y);
            Assert.Equal(5075640, element.Width);
            Assert.Equal(1439640, element.Height);
        }

        [Fact]
        public void CreateChart_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(x => x.GetId() == 4);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubGrFrame,
                ElementType = ShapeType.Chart
            };
            var mockPhService = Substitute.For<IPlaceholderService>();
            var creator = new ElementFactory(sldPart);
            var stubPhDic = new Dictionary<int, PlaceholderSL>();
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(ShapeType.Chart, element.Type);
            Assert.Equal(453241, element.X);
            Assert.Equal(752401, element.Y);
            Assert.Equal(2672732, element.Width);
            Assert.Equal(1819349, element.Height);
        }
    }
}
