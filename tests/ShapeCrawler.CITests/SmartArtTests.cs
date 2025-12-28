using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.CITests;

public class SmartArtTests : SCTest
{
    [Test]
    public void Nodes_AddNode_AddsTextToSmartArt()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var smartArt = pres.Slide(1).Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList).SmartArt;
        
        // Act
        var node1 = smartArt.Nodes.AddNode("Text 1");
        var node2 = smartArt.Nodes.AddNode("Text 2");
        var node3 = smartArt.Nodes.AddNode("Text 3");
        
        // Assert
        node1.Should().NotBeNull();
        node2.Should().NotBeNull();
        node3.Should().NotBeNull();
        
        smartArt.Nodes.Count.Should().Be(3);
        node1.Text.Should().Be("Text 1");
        node2.Text.Should().Be("Text 2");
        node3.Text.Should().Be("Text 3");
        
        ValidatePresentation(pres);
    }
    
    [Test]
    public void Node_Text_Setter_UpdatesNodeText()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList).SmartArt;
        var node = smartArt.Nodes.AddNode("Original Text");
        const string updatedText = "Updated Text";
        
        // Act
        node.Text = updatedText;
        
        // Assert
        node.Text.Should().Be(updatedText);
        
        ValidatePresentation(pres);
    }
}