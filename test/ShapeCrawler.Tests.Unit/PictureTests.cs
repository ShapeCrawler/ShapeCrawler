using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using ClosedXML;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Assert = Xunit.Assert;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Reliability", "CA2007:Consider calling ConfigureAwait on the awaited task")]
public class PictureTests : SCTest
{
    [Test]
    public async Task Image_BinaryData_returns_image_byte_array()
    {
        // Arrange
        var shapePicture1 = (IPicture)new SCPresentation(StreamOf("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);
        var shapePicture2 = (IPicture)new SCPresentation(StreamOf("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);

        // Act
        var shapePictureContentCase1 = shapePicture1.Image.BinaryData();
        var shapePictureContentCase2 =  shapePicture2.Image.BinaryData();

        // Assert
        shapePictureContentCase1.Should().NotBeEmpty();
        shapePictureContentCase2.Should().NotBeEmpty();
    }
        
    [Test]
    public async Task Image_GetBytes_returns_image_byte_array_of_Layout_picture()
    {
        // Arrange
        var pptxStream = StreamOf("pictures-case001.pptx");
        var presentation = new SCPresentation(pptxStream);
        var pictureShape = presentation.Slides[0].SlideLayout.Shapes.GetByName<IPicture>("Picture 7");
            
        // Act
        var picByteArray = pictureShape.Image.BinaryData();
            
        // Assert
        picByteArray.Should().NotBeEmpty();
    }
        
    [Test]
    public void Image_MIME_returns_MIME_type_of_image()
    {
        // Arrange
        var pptxStream = StreamOf("pictures-case001.pptx");
        var presentation = new SCPresentation(pptxStream);
        var image = presentation.Slides[0].SlideLayout.Shapes.GetByName<IPicture>("Picture 7").Image;
            
        // Act
        var mimeType = image.MIME;
            
        // Assert
        mimeType.Should().Be("image/png");
    }
        
    [Test]
    public void Image_GetBytes_returns_image_byte_array_of_Master_slide_picture()
    {
        // Arrange
        var pptxStream = StreamOf("pictures-case001.pptx");
        var presentation = new SCPresentation(pptxStream);
        var slideMaster = presentation.SlideMasters[0];
      var pictureShape = slideMaster.Shapes.GetByName<IPicture>("Picture 9");
            
        // Act
        var picByteArray = pictureShape.Image.BinaryData();
            
        // Assert
        picByteArray.Should().NotBeEmpty();
    }

    [Test]
    public void Image_SetImage_updates_picture_image()
    {
        // Arrange
        var pptx = StreamOf("009_table");
        var pngStream = StreamOf("test-image-2.png");
        var pres = new SCPresentation(pptx);
        var mStream = new MemoryStream();
        var picture = pres.Slides[1].Shapes.GetByName<IPicture>("Picture 1");
        var image = picture.Image!; 
        var lengthBefore = image.BinaryData().Length;
        
        // Act
        image.Update(pngStream);

        // Assert
        pres.SaveAs(mStream);
        pres = new SCPresentation(mStream);
        picture = pres.Slides[1].Shapes.GetByName<IPicture>("Picture 1");
        var lengthAfter = picture.Image!.BinaryData().Length;

        lengthAfter.Should().NotBe(lengthBefore);
    }

    [Test]
    public void Image_SvgContent_returns_svg_content()
    {
        // Arrange
        var pptxStream = StreamOf("pictures-case002.pptx");
        var pres = new SCPresentation(pptxStream);
        var picture = pres.Slides[0].Shapes.GetByName<IPicture>("Picture 1");

        // Act
        var svgContent = picture.SvgContent;
        
        // Assert
        svgContent.Should().NotBeEmpty();
    }

    [Test]
    public void Image_SetImage_should_not_update_image_of_other_grouped_picture()
    {
        // Arrange
        var pptx = StreamOf("pictures-case001.pptx");
        var image = GetTestBytes("test-image-2.png");
        var pres = new SCPresentation(pptx);
        var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 1");
        var groupedPicture1 = groupShape.Shapes.GetByName<IPicture>("Picture 1");
        var groupedPicture2 = groupShape.Shapes.GetByName<IPicture>("Picture 2");
        var stream = new MemoryStream();

        // Act
        groupedPicture1.Image.Update(image);

        // Assert
        pres.SaveAs(stream);
        var pictureContent1 = groupedPicture1.Image.BinaryData();
        var pictureContent2 = groupedPicture2.Image.BinaryData();
        pictureContent1.SequenceEqual(pictureContent2).Should().BeFalse();
    }
        
    [Test]
    public void Image_Name_Getter_returns_internal_image_file_name()
    {
        // Arrange
        var pptxStream = StreamOf("pictures-case001.pptx");
        var pres = new SCPresentation(pptxStream);
        var pictureImage = pres.Slides[0].Shapes.GetByName<IPicture>("Picture 3").Image;
            
        // Act
        var fileName = pictureImage.Name;
            
        // Assert
        fileName.Should().Be("image2.png");
    }

    [Test]
    public void Picture_DoNotParseStrangePicture_Test()
    {
        // TODO: Deeper learn such pictures, where content generated via a:ln
        // Arrange
        var pre = new SCPresentation(StreamOf("019.pptx"));

        // Act - Assert
        Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
    }
}