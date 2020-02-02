using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Extensions;
using SlideXML.Services.Placeholders;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represent a unit tests of <see cref="PlaceholderService"/> object.
    /// </summary>
    public class PlaceholderServiceTests
    {
        [Fact]
        public void Get_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>().Single(sp => sp.GetId() == 3);
            var sldLtPart = sldPart.SlideLayoutPart;
            var phService = new PlaceholderService(sldLtPart);

            // ACT
            var type = (P.PlaceholderValues)phService.TryGet(spId3).Type;

            // CLOSE
            xmlDoc.Close();
            ms.Dispose();

            // ASSERT
            Assert.Equal(P.PlaceholderValues.DateAndTime, type);
        }
    }
}
