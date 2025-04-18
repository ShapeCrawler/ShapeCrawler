using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

public class SmartArtTests : SCTest
{
    [Test]
    [Explicit("Should be fixed with https://github.com/ShapeCrawler/ShapeCrawler/issues/911")]
    public void AddSmartArt_CreatesBasicBlockList_WhenCalled()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0]; // First slide
        const int x = 50;
        const int y = 50;
        const int width = 400;
        const int height = 300;
        
        // Act
        var smartArt = slide.Shapes.AddSmartArt(x, y, width, height, SmartArtType.BasicBlockList);
        
        // Assert
        smartArt.Should().NotBeNull();
        smartArt.Should().BeOfType<SmartArt>();
        smartArt.X.Should().Be(x);
        smartArt.Y.Should().Be(y);
        smartArt.Width.Should().Be(width);
        smartArt.Height.Should().Be(height);
        
        pres.Validate();
    }
    
    [Test]
    public void SmartArtNodes_AddNode_AddsTextToSmartArt()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList);
        
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
        
        pres.Validate();
    }
    
    [Test]
    public void SmartArtNodes_ModifyNodeText_UpdatesNodeText()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList);
        var node = smartArt.Nodes.AddNode("Original Text");
        const string updatedText = "Updated Text";
        
        // Act
        node.Text = updatedText;
        
        // Assert
        node.Text.Should().Be(updatedText);
        
        pres.Validate();
    }
    
    [Test]
    public void SmartArtNodes_Enumeration_ReturnsAllNodes()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList);
        var expectedTexts = new[] { "Node 1", "Node 2", "Node 3" };
        
        foreach (var text in expectedTexts)
        {
            smartArt.Nodes.AddNode(text);
        }
        
        // Act
        var nodeTexts = smartArt.Nodes.Select(node => node.Text).ToArray();
        
        // Assert
        nodeTexts.Should().BeEquivalentTo(expectedTexts);
    }
}
