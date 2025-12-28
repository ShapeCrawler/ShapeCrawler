using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.DevTests;

public class SmartArtTests : SCTest
{
    [Test]
    public void AddSmartArt_AddsOneShapeInShapeCollection()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slide(1);
        var shapes = slide.Shapes;
        
        // Act
        shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList);
        
        // Assert
        shapes.Count.Should().Be(1);
    }
    
    [Test]
    public void SmartArtNodes_Enumeration_ReturnsAllNodes()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        var smartArt = slide.Shapes.AddSmartArt(50, 50, 400, 300, SmartArtType.BasicBlockList).SmartArt;
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
