using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal record ShapeFill : IShapeFill
{
    private readonly OpenXmlCompositeElement sdkTypedOpenXmlCompositeElement;
    private readonly OpenXmlPart sdkTypedOpenXmlPart;

    private SlidePictureImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private A.BlipFill? aBlipFill;

    internal ShapeFill(
        OpenXmlPart sdkTypedOpenXmlPart,
        OpenXmlCompositeElement sdkTypedOpenXmlCompositeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.sdkTypedOpenXmlCompositeElement = sdkTypedOpenXmlCompositeElement;
    }

    public string? Color
    {
        get
        {
            var aSolidFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return aRgbColorModelHex.Val!.ToString();
                }

                return this.ColorHexOrNullOf(aSolidFill.SchemeColor!.Val!);
            }

            return null;
        }
    }

    private string? ColorHexOrNullOf(string schemeColor)
    {
        var aColorScheme = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
        };

        var aColor2Type = aColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        var hex = aColor2Type?.RgbColorModelHex?.Val?.Value ?? aColor2Type?.SystemColor?.LastColor?.Value;

        if (hex != null)
        {
            return hex;
        }

        return null;
    }

    public double Alpha
    {
        get
        {
            const int defaultAlphaPercentages = 100;
            var aSolidFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    var alpha = aRgbColorModelHex.Elements<A.Alpha>().FirstOrDefault();
                    return alpha?.Val?.Value / 1000d ?? defaultAlphaPercentages;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.Alpha>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? defaultAlphaPercentages;
            }

            return defaultAlphaPercentages;
        }
    }

    public double LuminanceModulation
    {
        get
        {
            const double luminanceModulation = 100;
            var aSolidFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return luminanceModulation;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.LuminanceModulation>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? luminanceModulation;
            }

            return luminanceModulation;
        }
    }

    public double LuminanceOffset
    {
        get
        {
            const double defaultValue = 0;
            var aSolidFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (aSolidFill != null)
            {
                var aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return defaultValue;
                }

                var schemeColor = aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.LuminanceOffset>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? defaultValue;
            }

            return defaultValue;
        }
    }

    public IImage? Picture => this.GetPicture();

    public FillType Type => this.GetFillType();

    private FillType GetFillType()
    {
        var aSolidFillLocal = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (aSolidFillLocal != null)
        {
            return FillType.Solid;
        }

        var aGradFillLocal = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.GradientFill>();
        if (aGradFillLocal != null)
        {
            return FillType.Gradient;
        }

        var aBlipFillLocal = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.BlipFill>();
        if (aBlipFillLocal is not null)
        {
            return FillType.Picture;
        }

        var aPattFillLocal = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.PatternFill>();
        if (aPattFillLocal != null)
        {
            return FillType.Pattern;
        }

        if (this.sdkTypedOpenXmlCompositeElement.Ancestors<P.Shape>().First().UseBackgroundFill is not null)
        {
            return FillType.SlideBackground;
        }

        return FillType.NoFill;
    }

    public void SetPicture(Stream image)
    {
        if (this.Type == FillType.Picture)
        {
            this.pictureImage!.Update(image);
        }
        else
        {
            var rId = this.sdkTypedOpenXmlPart.AddImagePart(image);

            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.sdkTypedOpenXmlCompositeElement.Append(aBlipFill);

            this.aSolidFill?.Remove();
            this.aBlipFill = null;
            this.aGradFill?.Remove();
            this.aGradFill = null;
            this.aPattFill?.Remove();
            this.aPattFill = null;
        }
    }

    public void SetColor(string hex)
    {
        this.Initialize();
        this.sdkTypedOpenXmlCompositeElement.AddASolidFill(hex);
    }

    public void SetNoFill()
    {
        this.Initialize();
        this.sdkTypedOpenXmlCompositeElement.AddANoFill();
    }

    private void Initialize()
    {
        this.InitSolidFillOr();
    }

    private void InitSolidFillOr()
    {
        this.aSolidFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (this.aSolidFill == null)
        {
            this.aGradFill = this.sdkTypedOpenXmlCompositeElement!.GetFirstChild<A.GradientFill>();
            if (this.aGradFill == null)
            {
                this.InitPictureFillOr();
            }
        }
    }

    private void InitPictureFillOr()
    {
        this.aBlipFill = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var image = new SlidePictureImage(this.sdkTypedOpenXmlPart, this.aBlipFill.Blip!);
            this.pictureImage = image;
        }
    }

    private SlidePictureImage? GetPicture()
    {
        this.Initialize();

        return this.pictureImage;
    }
}