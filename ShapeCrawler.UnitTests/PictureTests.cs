using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using ClosedXML;
using FluentAssertions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tests.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests;

[SuppressMessage("Reliability", "CA2007:Consider calling ConfigureAwait on the awaited task")]
public class PictureTests : ShapeCrawlerTest
{
    [Fact]
    public async void Image_BinaryData_returns_image_byte_array()
    {
        // Arrange
        var pptx9 = GetTestStream("009_table.pptx");
        var pres9 = SCPresentation.Open(pptx9);
        var pptx18 = GetTestStream("018.pptx");
        var pres18 = SCPresentation.Open(pptx18);
        var shapePicture1 = (IPicture)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);
        var shapePicture2 = (IPicture)SCPresentation.Open(GetTestStream("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);

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
        var pptxStream = GetTestStream("009_table");
        var pngStream = GetTestStream("test-image-2.png");
        var pres = SCPresentation.Open(pptxStream);
        var mStream = new MemoryStream();
        var picture = pres.Slides[1].Shapes.GetByName<IPicture>("Picture 1");
        var image = picture.Image!; 
        var lengthBefore = image.BinaryData.Result.Length;
        
        // Act
        image.SetImage(pngStream);

        // Assert
        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        picture = pres.Slides[1].Shapes.GetByName<IPicture>("Picture 1");
        var lengthAfter = picture.Image!.BinaryData.Result.Length;

        lengthAfter.Should().NotBe(lengthBefore);
    }

    [Fact]
    public void Image_SvgContent_returns_svg_content()
    {
        // Arrange
        var pptxStream = GetTestStream("pictures-case002.pptx");
        using var pres = SCPresentation.Open(pptxStream);
        var picture = pres.Slides[0].Shapes.GetByName<IPicture>("Picture 1");

        // Act
        var svgContent = picture.SvgContent;
        
        // Assert
        svgContent.Should().NotBeEmpty();
    }

    [Fact]
    public async void Image_SetImage_should_not_update_image_of_other_grouped_picture()
    {
        // Arrange
        var pptxStream = GetTestStream("pictures-case001.pptx");
        var imgBytes = GetTestBytes("test-image-2.png");
        var pres = SCPresentation.Open(pptxStream);
        var groupShape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 1");
        var picture1 = groupShape.Shapes.GetByName<IPicture>("Picture 1");
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
        var pre = SCPresentation.Open(GetTestStream("019.pptx"));

        // Act - Assert
        Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
    }
}