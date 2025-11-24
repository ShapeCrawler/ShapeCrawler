using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.DevTests;

public class SmartArtTests : SCTest
{
    [Test]
    public void AddSmartArt_CreatesBasicBlockList_WhenCalled()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slide(1);
        const int x = 50;
        const int y = 50;
        const int width = 400;
        const int height = 300;
        
        // Act
        var smartArtShape = slide.Shapes.AddSmartArt(x, y, width, height, SmartArtType.BasicBlockList);
        
        // Assert
        ValidatePresentation(pres);
        smartArtShape.SmartArt.Should().NotBeNull();
        smartArtShape.X.Should().Be(x);
        smartArtShape.Y.Should().Be(y);
        smartArtShape.Width.Should().Be(width);
        smartArtShape.Height.Should().Be(height);
        var slidePart = pres.GetSDKPresentationDocument().PresentationPart!.SlideParts.First();
        var relationshipTypes = slidePart.Parts.Select(part => part.OpenXmlPart.RelationshipType).ToList();
        relationshipTypes.Should().Contain("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData");
    }
    
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
    public void SmartArtNodes_AddNode_AddsTextToSmartArt()
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
    public void SmartArtNodes_ModifyNodeText_UpdatesNodeText()
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
