using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.ShapeCollection;

internal sealed class SlideShapeOutline : IShapeOutline
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlCompositeElement sdkTypedOpenXmlCompositeElement;

    internal SlideShapeOutline(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlCompositeElement sdkTypedOpenXmlCompositeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.sdkTypedOpenXmlCompositeElement = sdkTypedOpenXmlCompositeElement;
    }

    public decimal Weight
    {
        get => this.ParseWeight();
        set => this.UpdateWeight(value);
    }

    public string? HexColor
    {
        get => this.ParseHexColor();
        set 
        {
            if (value == null)
            {
                this.Remove();                
            }
            else
            {
                this.UpdateHexColor(value);
            }
        }
    }

    private void UpdateWeight(decimal points)
    {
        var aOutline = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.sdkTypedOpenXmlCompositeElement.AddAOutline();
        }

        aOutline.Width = new Int32Value((Int32)UnitConverter.PointToEmu(points));
    }
    
    private void UpdateHexColor(string hex)
    {
        var aOutline = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>();

        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();
        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.sdkTypedOpenXmlCompositeElement.AddAOutline();
        }

        var aSolidFill = aOutline.GetFirstChild<A.SolidFill>();
        aNoFill?.Remove();
        aSolidFill?.Remove();

        var aSrgbColor = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill(aSrgbColor);
        aOutline.Append(aSolidFill);
    }

    private void Remove()
    {
        // Removing an outline really means ensuring that the shape has a 
        // 'NoFill' outline. If we just REMOVE the outline, then the style
        // will take over and give it the default outline, which is not
        // what we want
        A.Outline? outline = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>();
        if (outline is null)
        {
            outline = new A.Outline();
            this.sdkTypedOpenXmlCompositeElement.AppendChild(outline);
        }

        // Remove any explicit existing kinds of outline
        outline.RemoveAllChildren();

        // Add a nofill outline
        outline.AppendChild(new A.NoFill());
    }

    private decimal ParseWeight()
    {
        var width = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>()?.Width;
        if (width is null)
        {
            return 0;
        }

        var widthEmu = width.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }

    private string? ParseHexColor()
    {
        var aSolidFill = this.sdkTypedOpenXmlCompositeElement
            .GetFirstChild<A.Outline>()?
            .GetFirstChild<A.SolidFill>();
        if (aSolidFill is null)
        {
            return null;
        }

        var pSlideMaster = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
            _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).SlideMaster
        };
        var typeAndHex = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
        
        return typeAndHex.Item2;
    }
}