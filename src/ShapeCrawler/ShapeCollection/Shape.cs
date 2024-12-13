using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.ShapeCollection;

internal abstract class Shape : IShape
{
    protected readonly OpenXmlPart SdkTypedOpenXmlPart;
    protected readonly OpenXmlElement PShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private readonly ShapeId shapeId;

    internal Shape(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.SdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.PShapeTreeElement = pShapeTreeElement;
        this.position = new Position(sdkTypedOpenXmlPart, pShapeTreeElement);
        this.size = new ShapeSize(this.SdkTypedOpenXmlPart, pShapeTreeElement);
        this.shapeId = new ShapeId(pShapeTreeElement);
    }

    public virtual decimal X
    {
        get => this.position.X();
        set => this.position.UpdateX(value);
    }

    public virtual decimal Y
    {
        get => this.position.Y();
        set => this.position.UpdateY(value);
    }

    public decimal Width
    {
        get => this.size.Width();
        set => this.size.UpdateWidth(value);
    }

    public decimal Height
    {
        get => this.size.Height();
        set => this.size.UpdateHeight(value);
    }

    public int Id => this.shapeId.Value();

    public string Name => this.PShapeTreeElement.NonVisualDrawingProperties().Name!.Value!;
    
    public string AltText
    {
        get => this.PShapeTreeElement.NonVisualDrawingProperties().Description?.Value ?? string.Empty;
        set => this.PShapeTreeElement.NonVisualDrawingProperties().Description = new StringValue(value);
    }

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = this.PShapeTreeElement.NonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public bool IsPlaceholder => this.PShapeTreeElement.Descendants<P.PlaceholderShape>().Any();

    public PlaceholderType PlaceholderType
    {
        get
        {
            var pPlaceholderShape = this.PShapeTreeElement.Descendants<P.PlaceholderShape>().FirstOrDefault() ?? throw new SCException(
                    $"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.");
            var pPlaceholderValue = pPlaceholderShape.Type;
            if (pPlaceholderValue == null)
            {
                return PlaceholderType.Content;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Title)
            {
                return PlaceholderType.Title;
            }

            if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
            {
                return PlaceholderType.CenteredTitle;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Body)
            {
                return PlaceholderType.Text;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Diagram)
            {
                return PlaceholderType.SmartArt;
            }

            if (pPlaceholderValue == P.PlaceholderValues.ClipArt)
            {
                return PlaceholderType.OnlineImage;
            }

            var value = pPlaceholderValue.ToString() !;

            if (value == "dt")
            {
                return PlaceholderType.DateAndTime;
            }
 
            if (value == "ftr")
            {
                return PlaceholderType.Footer;
            }

            if (value == "sldNum")
            {
                return PlaceholderType.SlideNumber;
            }
            
            if(value == "pic")
            {
                return PlaceholderType.Picture;
            }
            
            if(value == "tbl")
            {
                return PlaceholderType.Table;
            }
            
            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), value, true);
        }
    } 
        
    public virtual Geometry GeometryType => Geometry.Rectangle;

    public decimal? CornerRoundedness
    {
        get
        {
            var spPr = this.PShapeTreeElement.Descendants<P.ShapeProperties>().First();
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();
            var shapeType = aPresetGeometry?.Preset?.Value;

            if (shapeType == A.ShapeTypeValues.RoundRectangle)
            {
                return GetRoundRectangleCornerRoundedness(aPresetGeometry!);
            }

            if (shapeType == A.ShapeTypeValues.Round2SameRectangle)
            {
                return GetTopRoundRectangleCornerRoundedness(aPresetGeometry!);
            }

            return null;
        }
        
        set
        {
            if (value is null)
            {
                throw new SCException("Not allowed to set null roundedness. Try 0 to straighten the corner.");
            }

            var spPr = this.PShapeTreeElement.Descendants<P.ShapeProperties>().First();
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();
            var shapeType = aPresetGeometry?.Preset?.Value;

            if (shapeType == A.ShapeTypeValues.RoundRectangle)
            {
                SetRoundRectangleCornerRoundedness(aPresetGeometry!, value ?? 0m);
            }

            if (shapeType == A.ShapeTypeValues.Round2SameRectangle)
            {
                SetTopRoundRectangleCornerRoundedness(aPresetGeometry!, value ?? 0m);
            }
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

            var elementText = regex.Match(this.PShapeTreeElement.InnerXml).Groups[1];
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
            this.PShapeTreeElement.InnerXml += customDataElement;
        }
    }

    public abstract ShapeType ShapeType { get; }
    
    public virtual bool HasOutline => false;

    public virtual IShapeOutline Outline => throw new SCException(
        $"Shape does not have outline. Use {nameof(IShape.HasOutline)} property to check if the shape has outline.");

    public virtual bool HasFill => false;

    public virtual IShapeFill Fill =>
        throw new SCException(
            $"Shape does not have fill. Use {nameof(IShape.HasFill)} property to check if the shape has fill.");

    public virtual bool IsTextHolder { get; protected init; }
    
    public virtual ITextBox TextBox { get; protected init; } = default(NullTextFrame);

    public virtual double Rotation
    {
        get
        {
            var pSpPr = this.PShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            var aTransform2D = pSpPr.Transform2D;
            if (aTransform2D == null)
            {
                aTransform2D = new ReferencedPShape(this.SdkTypedOpenXmlPart, this.PShapeTreeElement).ATransform2D();
                if (aTransform2D.Rotation is null)
                {
                    return 0;
                }

                return aTransform2D.Rotation.Value / 60000d; // OpenXML rotation angles are stored in units of 1/60,000th of a degree
            }
            
            return pSpPr.Transform2D!.Rotation!.Value / 60000d;
        }
    }

    public virtual bool Removeable => false;
    
    public string SdkXPath => new XmlPath(this.PShapeTreeElement).XPath;
    
    public OpenXmlElement SdkOpenXmlElement => this.PShapeTreeElement.CloneNode(true);

    public string Text
    {
        get => this.TextBox.Text;
        set => this.TextBox.Text = value;
    }

    public virtual ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public virtual IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media (audio, video, etc.");
    
    public virtual void Remove() => this.PShapeTreeElement.Remove();

    private static decimal? GetRoundRectangleCornerRoundedness(A.PresetGeometry aPresetGeometry)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.RoundRectangle)
        {
            return null;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj");
        if (sgs.Count() == 0)
        {
            // Has no shape guide. That means we're using the DEFAULT value, which is 0.35
            return 0.35m;
        }

        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        return GetCornerRoundednessFrom(sgs.Single());
    }

    private static void SetRoundRectangleCornerRoundedness(A.PresetGeometry aPresetGeometry, decimal value)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.RoundRectangle)
        {
            return;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj");
        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        // Will add a shape guide if there isn't already one
        var sg = sgs.SingleOrDefault()
            ?? avList.AppendChild(new A.ShapeGuide() { Name = "adj" }) 
            ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");

        sg.Formula = new StringValue($"val {(int)(value * 50000m)}");        
    }

    private static decimal? GetTopRoundRectangleCornerRoundedness(A.PresetGeometry aPresetGeometry)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.Round2SameRectangle)
        {
            return null;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>();
        var count = sgs.Count();
        if (count == 0)
        {
            // Has no shape guide. That means we're using the DEFAULT value, which is 0.35
            return 0.35m;
        }

        if (count != 2)
        {
            throw new SCException($"Malformed rounded rectangle. Expected 2 shape guides, found {count}. Please file a GitHub issue.");
        }

        var sg = sgs.Where(x => x.Name == "adj1").SingleOrDefault() ?? throw new SCException($"Malformed rounded rectangle. No shape guide named `adj1`. Please file a GitHub issue.");

        return GetCornerRoundednessFrom(sg);
    }

    private static void SetTopRoundRectangleCornerRoundedness(A.PresetGeometry aPresetGeometry, decimal value)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.Round2SameRectangle)
        {
            return;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj1");
        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        var sg = sgs.SingleOrDefault();
        if (sg is null)
        {
            // Has no adj1 shape guide. We need to add an adj1/adj2 pair
            sg = avList.AppendChild(new A.ShapeGuide() { Name = "adj1" }) ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");
            var _ = avList.AppendChild(new A.ShapeGuide() { Name = "adj2", Formula = "val 0" }) ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");
        }

        sg.Formula = new StringValue($"val {(int)(value * 50000m)}");        
    }

    private static decimal GetCornerRoundednessFrom(A.ShapeGuide sg)
    {
        var formula = sg.Formula?.Value ?? throw new SCException("Malformed rounded rectangle. Shape guide has no formula. Please file a GitHub issue.");

        var regex = new Regex("^val (?<value>[0-9]+)$");
        var match = regex.Match(formula);
        if (!match.Success)
        {
            throw new SCException("Malformed rounded rectangle. Formula has no value. Please file a GitHub issue.");
        }

        var value = int.Parse(match.Groups["value"].Value);

        return value / 50000m;
    }

}