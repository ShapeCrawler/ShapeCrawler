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
    IRoundedRectangle AddRoundedRectangle();
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
        var idAndName = this.GenerateIdAndName();

        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = A.ShapeTypeValues.Rectangle };
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

        var newPShape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)idAndName.Item1, Name = idAndName.Item2 },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                aParagraph));

        this.pShapeTree.Append(newPShape);
        
        var rectangle = new SCRectangle(this, newPShape, null);
     
        return rectangle;
    }
    
    public IRoundedRectangle AddRoundedRectangle()
    {
        // var idAndName = this.GenerateIdAndName();
        var id = 1;
        var name = "Rounded Rectangle 1";
        var x = 10;
        var y = 10;
        var width = 100;
        var height = 50;

        // p:nvSpPr
        var pNvSpPr = new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = (uint)id, Name = name },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties());
        
        // p:spPr
        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = A.ShapeTypeValues.RoundRectangle };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var widthEmu = UnitConverter.HorizontalPixelToEmu(width);
        var heightEmu = UnitConverter.VerticalPixelToEmu(height);
        shapeProperties.AddAXfrm(xEmu, yEmu, widthEmu, heightEmu);
        shapeProperties.Append(presetGeometry);

        // p:txBody
        var aRunProperties = new A.RunProperties { Language = "en-US" };
        var aText = new A.Text(string.Empty);
        var aRun = new A.Run(aRunProperties, aText);
        var aEndParaRPr = new A.EndParagraphRunProperties { Language = "en-US" };
        var aParagraph = new A.Paragraph(aRun, aEndParaRPr);

        var newPShape = new P.Shape(
            shapeProperties,
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                aParagraph));

        this.pShapeTree.Append(newPShape);

        return null!;
    }

    public IEnumerator<IAutoShape> GetEnumerator()
    {
        return this.autoShapes.OfType<IAutoShape>().GetEnumerator();
    }
    

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
    
    private (int, string) GenerateIdAndName()
    {
        var maxOrder = 0;
        var maxId = 0;
        foreach (var shape in this.allShapes)
        {
            if (shape.Id > maxId)
            {
                maxId = shape.Id;
            }

            var matchOrder = Regex.Match(shape.Name, "(?!AutoShape )\\d+");
            if (matchOrder.Success)
            {
                var order = int.Parse(matchOrder.Value);
                if (order > maxOrder)
                {
                    maxOrder = order;
                }
            }
        }

        var shapeId = maxId + 1;
        var shapeName = $"AutoShape {maxOrder + 1}";
        
        return (shapeId, shapeName);
    }
}