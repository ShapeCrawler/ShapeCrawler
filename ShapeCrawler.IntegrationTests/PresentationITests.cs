using FluentAssertions;
using ShapeCrawler.Tests.Shared;

namespace ShapeCrawler.IntegrationTests;

public class PresentationITests
{
    [Fact]
    public void Open_should_not_throw_exception()
    {
        // Arrange
        var originFilePath = Path.GetTempFileName();
        var savedAsFilePath = Path.GetTempFileName();
        var pptx = Assets.GetStream("001.pptx");
        File.WriteAllBytes(originFilePath, pptx.ToArray());
        var pres = SCPresentation.Open(originFilePath);
        pres.SaveAs(savedAsFilePath);
        pres.Close();

        // Act-Assert
        SCPresentation.Open(originFilePath);
        
        // Clean up
        pres.Close();
        File.Delete(originFilePath);
        File.Delete(savedAsFilePath);
    }
    
    [Fact]
    public void SaveAs_should_not_change_the_Original_Path_when_it_is_saved_to_New_Stream()
    {
        // Arrange
        var originalPath = Assets.GetPath("001.pptx");
        var pres = SCPresentation.Open(originalPath);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newStream = new MemoryStream();
        textBox.Text = originalText + "modified";

        // Act
        pres.SaveAs(newStream);

        // Assert
        pres.Close();
        pres = SCPresentation.Open(originalPath);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var autoShapeText = textBox!.Text;
        autoShapeText.Should().BeEquivalentTo(originalText);
        pres.Close();
            
        // Clean
        File.Delete(originalPath);
    }
           
    [Fact]
    public void SaveAs_should_not_change_the_Original_Stream_when_it_is_saved_to_New_Path()
    {
        // Arrange
        var originalFile = Assets.GetPath("001.pptx");
        var pres = SCPresentation.Open(originalFile);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newPath = Path.GetTempFileName();
        textBox.Text = originalText + "modified";

        // Act
        pres.SaveAs(newPath);

        // Assert
        pres.Close();
        pres = SCPresentation.Open(originalFile);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var autoShapeText = textBox!.Text;
        autoShapeText.Should().BeEquivalentTo(originalText);
        pres.Close();
            
        // Clean
        File.Delete(newPath);
    }
    
    [Fact]
    public void SaveAs_should_not_change_the_Original_Path_when_it_is_saved_to_New_Path()
    {
        // Arrange
        var originalPath = Assets.GetPath("001.pptx");
        var pres = SCPresentation.Open(originalPath);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newPath = Path.GetTempFileName();
        textBox.Text = originalText + "modified";

        // Act
        pres.SaveAs(newPath);

        // Assert
        pres.Close();
        pres = SCPresentation.Open(originalPath);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var autoShapeText = textBox!.Text; 
        autoShapeText.Should().BeEquivalentTo(originalText);
        pres.Close();
            
        // Clean
        File.Delete(originalPath);
        File.Delete(newPath);
    }
}