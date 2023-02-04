using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal abstract class SCShape : IShape
{
    protected SCShape(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
    {
        this.PShapeTreesChild = pShapeTreeChild;
        this.SlideBase = parentSlideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        this.GroupShape = parentShapeCollection.IsT1 ? parentShapeCollection.AsT1 : null;
        this.SlideObject = parentSlideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
    }

    internal event EventHandler<int>? XChanged;
    
    public int Id => (int)this.PShapeTreesChild.GetNonVisualDrawingProperties().Id!.Value;

    public string Name => this.PShapeTreesChild.GetNonVisualDrawingProperties().Name!;

    public bool Hidden =>
        this.DefineHidden(); // TODO: the Shape is inherited by LayoutShape, hence do we need this property?

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value ?? throw new ArgumentNullException(nameof(value)));
    }

    public abstract SCShapeType ShapeType { get; }

    public ISlideObject SlideObject { get; }

    public IPlaceholder? Placeholder => SCSlidePlaceholder.Create(this.PShapeTreesChild, this);

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
        get => this.GetHeightPixels();
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
            if (this.SlideBase is SCSlide slide)
            {
                return slide.SlideLayoutInternal.SlideMasterInternal;
            }

            if (this.SlideBase is SCSlideLayout layout)
            {
                return layout.SlideMasterInternal;
            }

            var master = (SCSlideMaster)this.SlideBase;
            return master;
        }
    }

    internal OpenXmlCompositeElement PShapeTreesChild { get; }

    internal SlideObject SlideBase { get; }

    private SCGroupShape? GroupShape { get; }

    internal abstract void Draw(SKCanvas canvas);
    
    protected virtual void SetXCoordinate(int xPx)
    {
        var pSpPr = this.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value;
            var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
            var yEmu = UnitConverter.VerticalPixelToEmu(referencedShape!.Y);
            var wEmu = UnitConverter.HorizontalEmuToPixel(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Offset!.X = UnitConverter.HorizontalPixelToEmu(xPx);
        }

        this.XChanged?.Invoke(this, this.X);
    }
    
    protected virtual void SetYCoordinate(int newYPixels)
    {
        var pSpPr = this.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;
        var aXfrm = pSpPr.Transform2D;
        if (aXfrm is null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            var referencedShape = placeholder.ReferencedShape.Value!;
            var xEmu = UnitConverter.HorizontalPixelToEmu(referencedShape.X);
            var yEmu = UnitConverter.HorizontalPixelToEmu(newYPixels);
            var wEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Width);
            var hEmu = UnitConverter.VerticalPixelToEmu(referencedShape.Height);
            pSpPr.AddAXfrm(xEmu, yEmu, wEmu, hEmu);
        }
        else
        {
            aXfrm.Offset!.Y = UnitConverter.HorizontalPixelToEmu(newYPixels);
        }
    }
    
    
    protected virtual void SetWidth(int newWPixels)
    {
        if (this.GroupShape is not null)
        {
            throw new RuntimeDefinedPropertyException("Width coordinate of grouped shape cannot be changed.");
        }
        
        var pSpPr = this.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;
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
        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().FirstOrDefault();
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
        if (this.GroupShape is not null)
        {
            throw new RuntimeDefinedPropertyException("Height coordinate of grouped shape cannot be changed.");
        }
        
        var pSpPr = this.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;
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
        this.PShapeTreesChild.InnerXml += customDataElement;
    }

    private string? GetCustomData()
    {
        var pattern = @$"<{SCConstants.CustomDataElementName}>(.*)<\/{SCConstants.CustomDataElementName}>";
        var regex = new Regex(pattern);
        var elementText = regex.Match(this.PShapeTreesChild.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    private bool DefineHidden()
    {
        var parsedHiddenValue = this.PShapeTreesChild.GetNonVisualDrawingProperties().Hidden?.Value;
        return parsedHiddenValue is true;
    }

    private int GetYCoordinate()
    {
        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!; 
            return placeholder.ReferencedShape.Value!.Y;
        }

        var yEmu = aOffset.Y!;

        if (this.GroupShape is not null)
        {
            var aTransformGroup =
                ((P.GroupShape)this.GroupShape.PShapeTreesChild).GroupShapeProperties!.TransformGroup!;
            yEmu = yEmu - aTransformGroup.ChildOffset!.Y! + aTransformGroup!.Offset!.Y!;
        }

        return UnitConverter.VerticalEmuToPixel(yEmu);
    }

    private int GetWidthPixels()
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!;
            return placeholder.ReferencedShape.Value!.Width;
        }

        return UnitConverter.HorizontalEmuToPixel(aExtents.Cx!);
    }

    private int GetHeightPixels()
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder!; 
            return placeholder.ReferencedShape.Value!.Height;
        }

        return UnitConverter.VerticalEmuToPixel(aExtents!.Cy!);
    }

    private SCGeometry GetGeometryType()
    {
        var spPr = this.PShapeTreesChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
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