using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal abstract class SCShapeFill : IShapeFill
{
    protected readonly TypedOpenXmlCompositeElement properties;
    protected BooleanValue? useBgFill;
    protected SCFillType fillType;
    protected A.NoFill? aNoFill;

    private const double DefaultAlphaPercentage = 100;
    private const double DefaultLuminanceModulationPercentage = 100;
    private const double DefaultLuminanceOffsetPercentage = 0;
    private readonly SlideStructure slideObject;
    private bool isDirty;
    private string? hexSolidColor;
    private double? alphaPercentage;
    private double? luminanceModulationPercentage;
    private double? luminanceOffsetPercentage;
    private SCImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private A.BlipFill? aBlipFill;

    internal SCShapeFill(SlideStructure slideObject, TypedOpenXmlCompositeElement properties)
    {
        this.slideObject = slideObject;
        this.properties = properties;
        this.isDirty = true;
    }

    public string? Color => this.GetHexSolidColor();

    public double? AlphaPercentage => this.GetAlphaPercentage();

    public double? LuminanceModulationPercentage => this.GetLuminanceModulationPercentage();

    public double? LuminanceOffsetPercentage => this.GetLuminanceOffsetPercentage();

    public IImage? Picture => this.GetPicture();

    public SCFillType Type => this.GetFillType();

    public void SetPicture(Stream imageStream)
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        if (this.Type == SCFillType.Picture)
        {
            this.pictureImage!.SetImage(imageStream);
        }
        else
        {
            var rId = this.slideObject.TypedOpenXmlPart.AddImagePart(imageStream);

            var aBlipFill = new DocumentFormat.OpenXml.Drawing.BlipFill();
            var aStretch = new DocumentFormat.OpenXml.Drawing.Stretch();
            aStretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
            aBlipFill.Append(new DocumentFormat.OpenXml.Drawing.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.properties.Append(aBlipFill);

            this.aSolidFill?.Remove();
            this.aBlipFill = null;
            this.aGradFill?.Remove();
            this.aGradFill = null;
            this.aPattFill?.Remove();
            this.aPattFill = null;
            this.useBgFill = false;
        }

        this.isDirty = true;
    }

    public void SetColor(string hex)
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        this.properties.AddASolidFill(hex);

        this.useBgFill = false;

        this.isDirty = true;
    }

    public void SetAlpha(double alphaPercentage)
    {
        // ToDo: Implement setting Alpha
        throw new NotImplementedException();
    }

    public void SetLuminanceModulation(double luminanceModulationPercentage)
    {
        // ToDo: Implement setting Luminance Modulation
        throw new NotImplementedException();
    }


    public void SetLuminanceOffset(double luminanceOffset)
    {
        // ToDo: Implement setting Luminance Offset
        throw new NotImplementedException();
    }

    protected virtual void InitSlideBackgroundFillOr()
    {
        this.aNoFill = this.properties.GetFirstChild<A.NoFill>();
        this.fillType = SCFillType.NoFill;
    }

    private SCFillType GetFillType()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.fillType;
    }

    private void Initialize()
    {
        this.InitSolidFillOr();
        this.isDirty = false;
    }

    private void InitSolidFillOr()
    {
        this.aSolidFill = this.properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
        if (this.aSolidFill != null)
        {
            var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                this.hexSolidColor = hexColor;
                this.alphaPercentage = this.GetAlphaPercentage(aRgbColorModelHex);
            }
            else
            {
                var hex = HexParser.FromSolidFill(this.aSolidFill, (SCSlideMaster)this.slideObject.SlideMaster);
                this.hexSolidColor = hex.Item2;

                var schemeColor = this.aSolidFill.SchemeColor!;
                this.alphaPercentage = this.GetAlphaPercentage(schemeColor);
                this.luminanceModulationPercentage = this.GetLuminanceModulationPercentage(schemeColor);
                this.luminanceOffsetPercentage = this.GetLuminanceOffsetPercentage(schemeColor);
            }

            this.fillType = SCFillType.Solid;
        }
        else
        {
            this.InitGradientFillOr();
        }
    }

    private void InitGradientFillOr()
    {
        this.aGradFill = this.properties!.GetFirstChild<A.GradientFill>();
        if (this.aGradFill != null)
        {
            this.fillType = SCFillType.Gradient;
        }
        else
        {
            this.InitPictureFillOr();
        }
    }

    private void InitPictureFillOr()
    {
        var xmlPart = this.slideObject.TypedOpenXmlPart;
        this.aBlipFill = this.properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var image = SCImage.ForAutoShapeFill(this.slideObject, xmlPart, this.aBlipFill);
            this.pictureImage = image;
            this.fillType = SCFillType.Picture;
        }
        else
        {
            this.InitPatternFillOr();
        }
    }

    private void InitPatternFillOr()
    {
        this.aPattFill = this.properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.PatternFill>();
        if (this.aPattFill != null)
        {
            this.fillType = SCFillType.Pattern;
        }
        else
        {
            this.InitSlideBackgroundFillOr();
        }
    }

    private string? GetHexSolidColor()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.hexSolidColor;
    }

    private double? GetAlphaPercentage()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.alphaPercentage;
    }

    private double? GetLuminanceModulationPercentage()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.luminanceModulationPercentage;
    }

    private double? GetLuminanceOffsetPercentage()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.luminanceOffsetPercentage;
    }

    private double GetAlphaPercentage(TypedOpenXmlCompositeElement element)
    {
        var alpha = element.Elements<A.Alpha>().FirstOrDefault();
        return alpha?.Val?.Value / 1000d ?? SCShapeFill.DefaultAlphaPercentage;
    }

    private double GetLuminanceModulationPercentage(TypedOpenXmlCompositeElement element)
    {
        var lumMod = element.Elements<A.LuminanceModulation>().FirstOrDefault();
        return lumMod?.Val?.Value / 1000d ?? SCShapeFill.DefaultLuminanceModulationPercentage;
    }

    private double GetLuminanceOffsetPercentage(TypedOpenXmlCompositeElement element)
    {
        var lumOff = element.Elements<A.LuminanceOffset>().FirstOrDefault();
        return lumOff?.Val?.Value / 1000d ?? SCShapeFill.DefaultLuminanceOffsetPercentage;
    }

    private SCImage? GetPicture()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.pictureImage;
    }
}