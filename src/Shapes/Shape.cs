using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal class Shape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement) : IShape
{
    public virtual decimal X
    {
        get => position.X;
        set => position.X = value;
    }

    public virtual decimal Y
    {
        get => position.Y;
        set => position.Y = value;
    }

    public virtual decimal Width
    {
        get => shapeSize.Width;
        set => shapeSize.Width = value;
    }

    public virtual decimal Height
    {
        get => shapeSize.Height;
        set => shapeSize.Height = value;
    }

    public IPresentation Presentation
    {
        get
        {
            var stream = new MemoryStream();
            new SCOpenXmlElement(pShapeTreeElement).PresentationDocument.Clone(stream);
            
            return new Presentation(stream);
        }
    }

    public int Id
    {
        get => shapeId.Value();
        internal set => shapeId.Update(value);
    }

    public string Name
    {
        get => pShapeTreeElement.NonVisualDrawingProperties().Name!.Value!;
        set => pShapeTreeElement.NonVisualDrawingProperties().Name = new StringValue(value);
    }

    public string AltText
    {
        get => pShapeTreeElement.NonVisualDrawingProperties().Description?.Value ?? string.Empty;
        set => pShapeTreeElement.NonVisualDrawingProperties().Description = new StringValue(value);
    }

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = pShapeTreeElement.NonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public PlaceholderType? PlaceholderType
    {
        get
        {
            var pPlaceholderShape = pShapeTreeElement.Descendants<P.PlaceholderShape>().FirstOrDefault();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            var pPlaceholderValue = pPlaceholderShape.Type;

            // Return default value if placeholder type is null
            if (pPlaceholderValue == null)
            {
                return ShapeCrawler.PlaceholderType.Content;
            }

            var placeholderValueMappings =
                new System.Collections.Generic.Dictionary<P.PlaceholderValues, PlaceholderType>
                {
                    { P.PlaceholderValues.Title, ShapeCrawler.PlaceholderType.Title },
                    { P.PlaceholderValues.Chart, ShapeCrawler.PlaceholderType.Chart },
                    { P.PlaceholderValues.CenteredTitle, ShapeCrawler.PlaceholderType.Title },
                    { P.PlaceholderValues.Body, ShapeCrawler.PlaceholderType.Text },
                    { P.PlaceholderValues.Diagram, ShapeCrawler.PlaceholderType.SmartArt },
                    { P.PlaceholderValues.ClipArt, ShapeCrawler.PlaceholderType.OnlineImage },
                };

            if (placeholderValueMappings.TryGetValue(pPlaceholderValue, out var mappedType))
            {
                return mappedType;
            }

            // Handle string-based values
            var value = pPlaceholderValue.ToString()!;
            var stringValueMappings =
                new System.Collections.Generic.Dictionary<string, PlaceholderType>(StringComparer.OrdinalIgnoreCase)
                {
                    { "dt", ShapeCrawler.PlaceholderType.DateAndTime },
                    { "ftr", ShapeCrawler.PlaceholderType.Footer },
                    { "sldNum", ShapeCrawler.PlaceholderType.SlideNumber },
                    { "pic", ShapeCrawler.PlaceholderType.Picture },
                    { "tbl", ShapeCrawler.PlaceholderType.Table },
                    { "sldImg", ShapeCrawler.PlaceholderType.SlideImage }
                };

            if (stringValueMappings.TryGetValue(value, out var stringMappedType))
            {
                return stringMappedType;
            }

            // Fallback for other values
            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), value, true);
        }
    }

    public virtual Geometry GeometryType
    {
        get
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).GeometryType;
        }

        set
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).GeometryType = value;
        }
    }

    public decimal CornerSize
    {
        get
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).CornerSize;
        }

        set
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).CornerSize = value;
        }
    }

    public decimal[] Adjustments
    {
        get
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).Adjustments;
        }

        set
        {
            var shapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).Adjustments = value;
        }
    }

    public string? CustomData
    {
        get
        {
            const string pattern = @$"<{"ctd"}>(.*)<\/{"ctd"}>";

#if NETSTANDARD2_0
            var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
            var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

            var elementText = regex.Match(pShapeTreeElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        set
        {
            var customDataElement =
                $@"<{"ctd"}>{value}</{"ctd"}>";
            pShapeTreeElement.InnerXml += customDataElement;
        }
    }

    public IShapeOutline Outline
    {
        get
        {
            var pShapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new SlideShapeOutline(pShapeProperties);
        }
    }

    public IShapeFill Fill
    {
        get
        {
            var pShapeProperties = pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeFill(pShapeProperties);
        }
    }

    public virtual ITextBox? TextBox => null;
    public virtual IPicture? Picture => null;
    public virtual IChart? Chart => null;
    public virtual ITable? Table => null;
    public virtual IOLEObject? OLEObject => null;
    public virtual IMedia? Media => null;
    public virtual ILine? Line => null;
    public virtual IShapeCollection? GroupedShapes => null;
    public ISmartArt? SmartArt => null;

    public virtual double Rotation
    {
        get
        {
            var pSpPr = pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            var aTransform2D = pSpPr.Transform2D;
            if (aTransform2D == null)
            {
                aTransform2D = new ReferencedPShape(pShapeTreeElement).ATransform2D();
                if (aTransform2D.Rotation is null)
                {
                    return 0;
                }

                return
                    aTransform2D.Rotation.Value /
                    60000d; // OpenXML rotation angles are stored in units of 1/60,000th of a degree
            }

            return pSpPr.Transform2D!.Rotation!.Value / 60000d;
        }
    }

    public bool Removable => false;

    public string SDKXPath => new XmlPath(pShapeTreeElement).XPath;

    public OpenXmlElement SDKOpenXmlElement => pShapeTreeElement.CloneNode(true);

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)pShapeTreeElement.Parent!;
        new SCPShapeTree(pShapeTree).Add(pShapeTreeElement);
    }

    public void Remove() => pShapeTreeElement.Remove();

    public virtual void CopyTo(P.ShapeTree pShapeTree) => new SCPShapeTree(pShapeTree).Add(pShapeTreeElement);
    
    public virtual void SetText(string text) => throw new SCException("The shape doesn't contain text content");

    public virtual void SetMarkdownText(string text) => throw new SCException("The shape doesn't contain text content");
    
    public virtual void SetImage(string imagePath) => throw new SCException();
    public virtual void SetFontName(string fontName) => throw new SCException("The shape doesn't contain text content");

    public virtual void SetFontSize(decimal fontSize) => throw new SCException("The shape doesn't contain text content");

    public virtual void SetFontColor(string colorHex) => throw new SCException("The shape doesn't contain text content");

    public virtual void SetVideo(Stream video) => throw new SCException("The shape doesn't support video content");
    public IShape GroupedShape(string name)
    {
        if (this.GroupedShapes == null)
        {
            throw new SCException("The current shape is not a group shape.");
        }

        var groupedShape = this.GroupedShapes.FirstOrDefault(shape => shape.Name == name);
        if (groupedShape == null)
        {
            throw new SCException($"Grouped shape with name '{name}' not found.");
        }

        return groupedShape;
    }
}