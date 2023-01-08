using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections;

/// <summary>
///     Represents a collection of AutoShapes.
/// </summary>
public interface IAutoShapeCollection : IReadOnlyList<IAutoShape>
{
    /// <summary>
    ///     Adds a new Rectangle shape.
    /// </summary>
    IRectangle AddRectangle(int x, int y, int w, int h);
    
    /// <summary>
    ///     Adds a new Rounded Rectangle shape. 
    /// </summary>
    IRoundedRectangle AddRoundedRectangle(int x, int y, int w, int h);
}

internal class AutoShapeCollection : IAutoShapeCollection
{
    private readonly P.ShapeTree pShapeTree;
    private readonly IAutoShape[] autoShapes;
    private readonly IEnumerable<IShape> allShapes;

    internal AutoShapeCollection(IEnumerable<IShape> allShapes, P.ShapeTree pShapeTree, ShapeCollection parentShapeCollection)
    {
        this.allShapes = allShapes;
        this.pShapeTree = pShapeTree;
        this.ParentShapeCollection = parentShapeCollection;
        this.autoShapes = allShapes.Where(shape => shape is AutoShape).OfType<IAutoShape>().ToArray();
    }
    
    public int Count => this.autoShapes.Length;

    internal ShapeCollection ParentShapeCollection { get; }
    
    public IAutoShape this[int index] => this.autoShapes[index];
    
    public IRectangle AddRectangle(int x, int y, int width, int height)
    {
        var newPShape = this.CreatePShape(x, y, width, height, A.ShapeTypeValues.Rectangle);

        this.pShapeTree.Append(newPShape);
        
        var rectangle = new SCRectangle(this, newPShape, null);
        rectangle.Outline.Color = "000000";
     
        return rectangle;
    }

    public IRoundedRectangle AddRoundedRectangle(int x, int y, int width, int height)
    {
        var newPShape = this.CreatePShape(x, y, width, height, A.ShapeTypeValues.RoundRectangle);

        var roundedRectangle = new SCRoundedRectangle(this, newPShape, null);

        roundedRectangle.Outline.Color = "000000";

        return roundedRectangle;
    }

    public IEnumerator<IAutoShape> GetEnumerator()
    {
        return this.autoShapes.OfType<IAutoShape>().GetEnumerator();
    }
    

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
    
    private P.Shape CreatePShape(int x, int y, int width, int height, A.ShapeTypeValues form)
    {
        var idAndName = this.GenerateIdAndName();
        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = form };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var widthEmu = UnitConverter.HorizontalPixelToEmu(width);
        var heightEmu = UnitConverter.VerticalPixelToEmu(height);
        shapeProperties.AddAXfrm(xEmu, yEmu, widthEmu, heightEmu);
        shapeProperties.Append(presetGeometry);

        var aRunProperties = new A.RunProperties { Language = "en-US" };
        var aText = new A.Text(string.Empty);
        var aRun = new A.Run(aRunProperties, aText);
        var aEndParaRPr = new A.EndParagraphRunProperties { Language = "en-US" };
        var aParagraph = new A.Paragraph(aRun, aEndParaRPr);

        var pShape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)idAndName.Item1, Name = idAndName.Item2 },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                aParagraph));

        return pShape;
    }
    
    private (int, string) GenerateIdAndName()
    {
        var maxId = this.allShapes.Max(s => s.Id);
        var maxOrder = Regex.Matches(string.Join(string.Empty, this.allShapes.Select(s => s.Name)), "\\d+")
            #if NETSTANDARD2_0
            .Cast<Match>()
            #endif
            .Select(m => int.Parse(m.Value))
            .DefaultIfEmpty(0)
            .Max();
        
        return (maxId + 1, $"AutoShape {maxOrder + 1}");
    }
}