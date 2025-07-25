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

internal sealed class Shape(Position position, ShapeSize size, ShapeId shapeId, OpenXmlElement pShapeTreeElement)
    : IShape
{
    public decimal X
    {
        get => position.X;
        set => position.X = value;
    }

    public decimal Y
    {
        get => position.Y;
        set => position.Y = value;
    }

    public decimal Width
    {
        get => size.Width;
        set => size.Width = value;
    }

    public decimal Height
    {
        get => size.Height;
        set => size.Height = value;
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

    public Geometry GeometryType
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

    public ShapeContent ShapeContent => ShapeContent.Shape;

    public IShapeOutline Outline
    {
        get
        {
            var pShapeProperties = pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            return new SlideShapeOutline(pShapeProperties);
        }
    }

    public IShapeFill Fill
    {
        get
        {
            var pShapeProperties = pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            return new ShapeFill(pShapeProperties);
        }
    }

    public ITextBox? TextBox => null;

    public double Rotation
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

    public ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeContent)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a media (audio, video, etc.");

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)pShapeTreeElement.Parent!;
        new SCPShapeTree(pShapeTree).Add(pShapeTreeElement);
    }

    public void SetText(string text) => throw new SCException(
        $"The shape is not a text shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a text shape.");

    public void SetImage(string imagePath) => throw new SCException(
        $"The shape is not an image shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is an image shape.");

    public void Remove() => pShapeTreeElement.Remove();

    public void CopyTo(P.ShapeTree pShapeTree) => new SCPShapeTree(pShapeTree).Add(pShapeTreeElement);

    public void SetFontName(string fontName) => throw new SCException(
        $"The shape is not a text shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a text shape.");

    public void SetFontSize(decimal fontSize) => throw new SCException(
        $"The shape is not a text shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a text shape.");

    public void SetFontColor(string colorHex) => throw new SCException(
        $"The shape is not a text shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a text shape.");

    public void SetVideo(Stream video)
    {
        throw new NotImplementedException();
    }
}