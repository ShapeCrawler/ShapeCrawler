using PptxXML.Models;
using PptxXML.Tests.Helpers;
using System.Linq;
using NSubstitute;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;
using PptxXML.Services;
using PptxXML.Services.Builders;
using PptxXML.Services.Placeholder;
using PptxXML.Services.Placeholders;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represents unit tests for the <see cref="SlideCollection"/> class.
    /// </summary>
    public class SlideCollectionTests
    {
        [Fact]
        public void Add_AddedOneItem_SlidesNumberIsOne()
        {
            // ARRANGE
            var xmlDoc = DocHelper.Open(Properties.Resources._001);
            var slides = new SlideCollection(xmlDoc);
            var mockTxtBuilder = Substitute.For<ITextBodyExBuilder>();
            var elementCreator = new ElementFactory(new ShapeEx.Builder(new BackgroundImageFactory(), mockTxtBuilder));
            var treeParser = new GroupShapeTypeParser();
            var builder = new GroupEx.Builder(treeParser, elementCreator);
            var bgImgFactory = new BackgroundImageFactory();
            var mockPreSettings = Substitute.For<IPreSettings>();
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var newSlide = new SlideEx(sldPart, 1, elementCreator, treeParser, builder, new SlideLayoutPartParser(), bgImgFactory, mockPreSettings);

            // ACT
            slides.Add(newSlide);

            // CLEAN
            xmlDoc.Dispose();

            // ASSERT
            Assert.Single(slides);
        }

        /// <State>
        /// - there is presentation with two slides;
        /// - first slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation contains single slide with 1 number.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test1()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._007_2_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Equal(1, slides.Single().Number);

            // CLEAN
            pre.Dispose();
        }

        /// <State>
        /// - there is a single slide presentation;
        /// - slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation is empty.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test2()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Empty(slides);

            // CLEAN
            pre.Dispose();
        }
    }
}
