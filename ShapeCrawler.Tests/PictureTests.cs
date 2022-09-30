using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using ClosedXML;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests
{
    [SuppressMessage("Reliability", "CA2007:Consider calling ConfigureAwait on the awaited task")]
    public class PictureTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public PictureTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public async void Image_BinaryData_returns_image_byte_array()
        {
            // Arrange
            var shapePicture1 = (IPicture)_fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 3);
            var shapePicture2 = (IPicture)_fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 7);

            // Act
            var shapePictureContentCase1 = await shapePicture1.Image.BinaryData.ConfigureAwait(false);
            var shapePictureContentCase2 = await shapePicture2.Image.BinaryData.ConfigureAwait(false);

            // Assert
            shapePictureContentCase1.Should().NotBeEmpty();
            shapePictureContentCase2.Should().NotBeEmpty();
        }
        
        [Fact]
        public async void Image_GetBytes_returns_image_byte_array_of_Layout_picture()
        {
            // Arrange
            var pptxStream = GetTestStream("pictures-case001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var pictureShape = presentation.Slides[0].SlideLayout.Shapes.GetByName<IPicture>("Picture 7");
            
            // Act
            var picByteArray = await pictureShape.Image.BinaryData;
            
            // Assert
            picByteArray.Should().NotBeEmpty();
        }
        
        [Fact]
        public void Image_MIME_returns_MIME_type_of_image()
        {
            // Arrange
            var pptxStream = GetTestStream("pictures-case001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var image = presentation.Slides[0].SlideLayout.Shapes.GetByName<IPicture>("Picture 7").Image;
            
            // Act
            var mimeType = image.MIME;
            
            // Assert
            mimeType.Should().Be("image/png");
        }
        
        [Fact]
        public void Image_GetBytes_returns_image_byte_array_of_Master_slide_picture()
        {
            // Arrange
            var pptxStream = GetTestStream("pictures-case001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var slideMaster = presentation.SlideMasters[0];
            var pictureShape = slideMaster.Shapes.GetByName<IPicture>("Picture 9");
            
            // Act
            var picByteArray = pictureShape.Image.BinaryData.Result;
            
            // Assert
            picByteArray.Should().NotBeEmpty();
        }

        [Fact]
        public void Image_SetImage_updates_picture_image()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Properties.Resources._009);
            MemoryStream imageStream = new (TestFiles.Images.imageByteArray02);
            MemoryStream preStream = new();
            IPicture picture = (IPicture) presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            int lengthBefore = picture.Image.BinaryData.Result.Length;

            // Act
            picture.Image.SetImage(imageStream);

            // Assert
            presentation.SaveAs(preStream);
            presentation.Close();
            presentation = SCPresentation.Open(preStream);
            picture = (IPicture)presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            int lengthAfter = picture.Image.BinaryData.Result.Length;

            lengthAfter.Should().NotBe(lengthBefore);
        }

        [Fact]
        public async void Image_SetImage_should_not_update_image_of_other_grouped_picture()
        {
            // Arrange
            var pptxStream = GetTestStream("pictures-case001.pptx");
            var imgBytes = GetTestBytes("test-image-2.png");
            var pres = SCPresentation.Open(pptxStream);
            var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 1");
            var picture1 = groupShape.Shapes.OfType<IPicture>().First(s => s.Name == "Picture 1");
            var picture2 = groupShape.Shapes.OfType<IPicture>().First(s => s.Name == "Picture 2");
            var mStream = new MemoryStream();

            // Act
            picture1.Image.SetImage(imgBytes);

            // Assert
            pres.SaveAs(mStream);
            var bytes1 = await picture1.Image.BinaryData; 
            var bytes2 = await picture2.Image.BinaryData;
            bytes1.SequenceEqual(bytes2).Should().BeFalse();
        }
        
        [Fact]
        public void Image_Name_Getter_returns_internal_image_file_name()
        {
            // Arrange
            var pptxStream = GetTestStream("pictures-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var pictureImage = pres.Slides[0].Shapes.GetByName<IPicture>("Picture 3").Image;
            
            // Act
            var fileName = pictureImage.Name;
            
            // Assert
            fileName.Should().Be("image2.png");
        }

        [Fact]
        public void Picture_DoNotParseStrangePicture_Test()
        {
            // TODO: Deeper learn such pictures, where content generated via a:ln
            // Arrange
            var pre = _fixture.Pre019;

            // Act - Assert
            Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
        }
    }
}
