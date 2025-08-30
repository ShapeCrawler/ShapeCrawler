using NUnit.Framework;

namespace ShapeCrawler.DevTests;

public class TableShapeTests
{
    [Test]
    public void Width()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.Table("Table 1", 100, 100, 2, 1);
            });
        });
        var tableShape = pres.Slide(1).Shape("Table 1");
        var newWidth = tableShape.Width *=1.25m;
        
        // Act
        tableShape.Width *= newWidth;  
        pres.Save(@"c:\temp\output.pptx");
        
        // Assert
        pres.Validate();
    }
}