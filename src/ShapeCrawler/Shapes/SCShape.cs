using System;
using System.Linq;
using System.Text.RegularExpressions;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal abstract class SCShape : IShape
{
    internal OneOf<ShapeCollection, SCGroupShape> shapeCollectionOf;
    internal OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf;

    protected SCShape(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollectionStructureOf)
    {
        this.PShapeTreeChild = pShapeTreeChild;
        this.slideOf = slideOf;
        this.shapeCollectionOf = parentShapeCollectionStructureOf;
        this.SlideStructure = slideOf.Match(slide => slide as SlideStructure, layout => layout, master => master);
        this.GroupShape = parentShapeCollectionStructureOf.IsT1 ? parentShapeCollectionStructureOf.AsT1 : null;
    }
    
    internal event EventHandler<int>? XChanged;

    internal event EventHandler<int>? YChanged;

    public int Id => (int)this.PShapeTreeChild.GetNonVisualDrawingProperties().Id!.Value;

    public string Name => this.PShapeTreeChild.GetNonVisualDrawingProperties().Name!;

    public bool Hidden =>
        this.DefineHidden(); // TODO: the Shape is inherited by LayoutShape, hence do we need this property?

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value ?? throw new ArgumentNullException(nameof(value)));
    }

    public abstract SCShapeType ShapeType { get; }
    
    public ISlideStructure SlideStructure { get; }
    
    public IPlaceholder? Placeholder => SCSlidePlaceholder.Create(this.PShapeTreeChild, this);

    public virtual SCGeometry GeometryType => this.GetGeometryType();

    public int X
    {
        get => this.GetXCoordinate();
        set => this.SetXCoordinate(value);
    }

    public int Y
    {
        get => this.GetYCoordinate();
        set => this.SetYCoordinate(value);
    }

    public int Height
    {
        get => this.GetHeight();
        set => this.SetHeight(value);
    }

    public int Width
    {
        get => this.GetWidthPixels();
        set => this.SetWidth(value);
    }
    
    internal SCSlideMaster SlideMasterInternal
    {
        get
        {
            if (this.SlideStructure is SCSlide slide)
            {
                return slide.SlideLayoutInternal.SlideMasterInternal;
            }

            if (this.SlideStructure is SCSlideLayout layout)
            {
                return layout.SlideMasterInternal;
            }

            var master = (SCSlideMaster)this.SlideStructure;
            return master;
        }
    }

    internal OpenXmlCompositeElement PShapeTreeChild { get; }

    private SCGroupShape? GroupShape { get; }

    public IAutoShape? AsAutoShape()
    {
        return this as IAutoShape;
    }
    
    internal abstract void Draw(SKCanvas canvas);
    
    internal abstract string ToJson();
    
    internal abstract IHtmlElement ToHtmlElement();
    
    protected virtual void SetXCoordinate(int newXPx)
    {
        var pSpPr = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(newXPx);
            var yEmu = UnitConverter.VerticalPixelToEmu(referencedShape!.Y);
            var wEmu = UnitConverter.HorizontalEmuToPixel(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Offset!.X = UnitConverter.HorizontalPixelToEmu(newXPx);
        }

        this.XChanged?.Invoke(this, this.X);
    }
    
    protected virtual void SetYCoordinate(int newYPx)
    {
        var pSpPr = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value!;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(newYPx);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Offset!.Y = UnitConverter.HorizontalPixelToEmu(newYPx);
        }
        
        this.YChanged?.Invoke(this, this.Y);
    }
    
    
    protected virtual void SetWidth(int newWPixels)
    {
        if (this.GroupShape is not null)
        {
            throw new RuntimeDefinedPropertyException("Width coordinate of grouped shape cannot be changed.");
        }
        
        var pSpPr = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape!.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.Y);
            var wEmu = UnitConverter.VerticalPixelToEmu(newWPixels);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Extents!.Cx = UnitConverter.HorizontalPixelToEmu(newWPixels);
        }
    }
    
    protected virtual int GetXCoordinate()
    {
        var aOffset = this.PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value; 
            
            return referencedShape!.X;
        }

        var xEmu = aOffset.X!.Value;
        if (this.GroupShape == null)
        {
            return UnitConverter.HorizontalEmuToPixel(xEmu);    
        }
        
        var groupedShapeX = aOffset.X!.Value;
        var groupShapeX = this.GroupShape!.ATransformGroup.Offset!.X!.Value;
        var groupShapeChildX = this.GroupShape!.ATransformGroup.ChildOffset!.X!.Value;
        var absoluteX = groupShapeX - (groupShapeChildX - groupedShapeX);

        return UnitConverter.HorizontalEmuToPixel(absoluteX);
    }
    
    protected virtual void SetHeight(int newHPixels)
    {
        var pSpPr = this.PShapeTreeChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape!.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.Y);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(newHPixels);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Extents!.Cy = UnitConverter.HorizontalPixelToEmu(newHPixels);
        }
    }
    
    private void SetCustomData(string value)
    {
        string customDataElement =
            $@"<{SCConstants.CustomDataElementName}>{value}</{SCConstants.CustomDataElementName}>";
        this.PShapeTreeChild.InnerXml += customDataElement;
    }

    private string? GetCustomData()
    {
        var pattern = @$"<{SCConstants.CustomDataElementName}>(.*)<\/{SCConstants.CustomDataElementName}>";
        var regex = new Regex(pattern);
        var elementText = regex.Match(this.PShapeTreeChild.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    private bool DefineHidden()
    {
        var parsedHiddenValue = this.PShapeTreeChild.GetNonVisualDrawingProperties().Hidden?.Value;
        return parsedHiddenValue is true;
    }

    private int GetYCoordinate()
    {
        var aOffset = this.PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!; 
            return placeholder.ReferencedShape.Value!.Y;
        }

        var yEmu = aOffset.Y!;

        if (this.GroupShape is not null)
        {
            var aTransformGroup =
                ((P.GroupShape)this.GroupShape.PShapeTreeChild).GroupShapeProperties!.TransformGroup!;
            yEmu = yEmu - aTransformGroup.ChildOffset!.Y! + aTransformGroup!.Offset!.Y!;
        }

        return UnitConverter.VerticalEmuToPixel(yEmu);
    }

    private int GetWidthPixels()
    {
        var aExtents = this.PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            return placeholder.ReferencedShape.Value!.Width;
        }

        return UnitConverter.HorizontalEmuToPixel(aExtents.Cx!);
    }

    private int GetHeight()
    {
        var aExtents = this.PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!; 
            return placeholder.ReferencedShape.Value!.Height;
        }

        return UnitConverter.VerticalEmuToPixel(aExtents!.Cy!);
    }

    private SCGeometry GetGeometryType()
    {
        var spPr = this.PShapeTreeChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
        var aTransform2D = spPr.Transform2D;
        if (aTransform2D != null)
        {
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

            // Placeholder can have transform on the slide, without having geometry
            if (aPresetGeometry == null)
            {
                if (spPr.OfType<A.CustomGeometry>().Any())
                {
                    return SCGeometry.Custom;
                }
            }
            else
            {
                var name = aPresetGeometry.Preset!.Value.ToString();
                Enum.TryParse(name, true, out SCGeometry geometryType);
                return geometryType;
            }
        }

        var placeholder = this.Placeholder as SCPlaceholder;
        if (placeholder?.ReferencedShape.Value != null)
        {
            return placeholder.ReferencedShape.Value.GeometryType;
        }

        return SCGeometry.Rectangle; // return default
    }
}