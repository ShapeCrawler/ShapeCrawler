using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFill : IShapeFill
{
    private readonly OpenXmlCompositeElement openXmlCompositeElement;
    private readonly OpenXmlPart openXmlPart;

    private SlidePictureImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPatternFill;
    private A.BlipFill? aBlipFill;

    internal ShapeFill(OpenXmlPart openXmlPart, OpenXmlCompositeElement openXmlCompositeElement)
    {
        this.openXmlPart = openXmlPart;
        this.openXmlCompositeElement = openXmlCompositeElement;
    }

    public string? Color
    {
        get
        {
            var aSolidFill = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
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
        var aColorScheme = this.openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)this.openXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
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
            var aSolidFill = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
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
            var aSolidFill = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
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
            var aSolidFill = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
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

    public IImage? Picture => this.GetPictureImage();

    public FillType Type => this.GetFillType();

    private FillType GetFillType()
    {
        var aSolidFillLocal = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (aSolidFillLocal != null)
        {
            return FillType.Solid;
        }

        var aGradFillLocal = this.openXmlCompositeElement.GetFirstChild<A.GradientFill>();
        if (aGradFillLocal != null)
        {
            return FillType.Gradient;
        }

        var aBlipFillLocal = this.openXmlCompositeElement.GetFirstChild<A.BlipFill>();
        if (aBlipFillLocal is not null)
        {
            return FillType.Picture;
        }

        var aPattFillLocal = this.openXmlCompositeElement.GetFirstChild<A.PatternFill>();
        if (aPattFillLocal != null)
        {
            return FillType.Pattern;
        }

        if (this.openXmlCompositeElement.Ancestors<P.Shape>().FirstOrDefault()?.UseBackgroundFill is not null)
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
            (var rId, _) = this.openXmlPart.AddImagePart(image, "image/png");

            // This could be refactored to DRY vs SlideShapes.CreatePPicture.
            // In the process, the image could be de-duped also.
            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.openXmlCompositeElement.Append(aBlipFill);

            this.aSolidFill?.Remove();
            this.aBlipFill = null;
            this.aGradFill?.Remove();
            this.aGradFill = null;
            this.aPatternFill?.Remove();
            this.aPatternFill = null;
        }
    }

    public void SetColor(string hex)
    {
        this.InitSolidFillOr();
        this.openXmlCompositeElement.AddSolidFill(hex);
    }

    public void SetNoFill()
    {
        this.InitSolidFillOr();
        this.openXmlCompositeElement.AddNoFill();
    }

    private void InitSolidFillOr()
    {
        this.aSolidFill = this.openXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (this.aSolidFill == null)
        {
            this.aGradFill = this.openXmlCompositeElement!.GetFirstChild<A.GradientFill>();
            if (this.aGradFill == null)
            {
                this.InitPictureFillOr();
            }
        }
    }

    private void InitPictureFillOr()
    {
        this.aBlipFill = this.openXmlCompositeElement.GetFirstChild<A.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var image = new SlidePictureImage(this.aBlipFill.Blip!);
            this.pictureImage = image;
        }
    }

    private SlidePictureImage? GetPictureImage()
    {
        this.InitSolidFillOr();

        return this.pictureImage;
    }
}