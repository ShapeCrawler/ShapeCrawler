using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFill(OpenXmlCompositeElement openXmlCompositeElement) : IShapeFill
{
    private SlidePictureImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.BlipFill? aBlipFill;

    public string? Color
    {
        get
        {
            this.aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (this.aSolidFill != null)
            {
                var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return aRgbColorModelHex.Val!.ToString();
                }

                return this.ColorHexOrNullOf(this.aSolidFill.SchemeColor!.Val!);
            }

            return null;
        }
    }

    public double Alpha
    {
        get
        {
            const int defaultAlphaPercentages = 100;
            this.aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (this.aSolidFill != null)
            {
                var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    var alpha = aRgbColorModelHex.Elements<A.Alpha>().FirstOrDefault();
                    return alpha?.Val?.Value / 1000d ?? defaultAlphaPercentages;
                }

                var schemeColor = this.aSolidFill.SchemeColor!;
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
            this.aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (this.aSolidFill != null)
            {
                var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return luminanceModulation;
                }

                var schemeColor = this.aSolidFill.SchemeColor!;
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
            this.aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
            if (this.aSolidFill != null)
            {
                var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
                if (aRgbColorModelHex != null)
                {
                    return defaultValue;
                }

                var schemeColor = this.aSolidFill.SchemeColor!;
                var schemeAlpha = schemeColor.Elements<A.LuminanceOffset>().FirstOrDefault();
                return schemeAlpha?.Val?.Value / 1000d ?? defaultValue;
            }

            return defaultValue;
        }
    }

    public IImage? Picture => this.GetPictureImage();

    public FillType Type => this.GetFillType();

    public void SetPicture(Stream image)
    {
        var openXmlPart = openXmlCompositeElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (this.Type == FillType.Picture)
        {
            this.pictureImage!.Update(image);
        }
        else
        {
            openXmlCompositeElement.GetFirstChild<A.SolidFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.GradientFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.PatternFill>()?.Remove();
            openXmlCompositeElement.GetFirstChild<A.NoFill>()?.Remove();

            (var rId, _) = openXmlPart.AddImagePart(image, "image/png");

            this.aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            this.aBlipFill.Append(new A.Blip { Embed = rId });
            this.aBlipFill.Append(aStretch);

            var aOutline = openXmlCompositeElement.GetFirstChild<A.Outline>();
            if (aOutline != null)
            {
                openXmlCompositeElement.InsertBefore(this.aBlipFill, aOutline);
            }
            else
            {
                openXmlCompositeElement.Append(this.aBlipFill);
            }

            this.aSolidFill = null;
            this.aGradFill = null;
            this.pictureImage = new SlidePictureImage(this.aBlipFill.Blip!);
        }
    }

    public void SetColor(string hex)
    {
        this.InitSolidFillOr();
        openXmlCompositeElement.AddSolidFill(hex);
    }

    public void SetNoFill()
    {
        this.InitSolidFillOr();
        openXmlCompositeElement.AddNoFill();
    }

    private static A.ColorScheme GetColorScheme(OpenXmlPart openXmlPart)
    {
        return openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .ColorScheme!,
            _ => ((SlideMasterPart)openXmlPart).ThemePart!.Theme.ThemeElements!.ColorScheme!
        };
    }

    private void InitSolidFillOr()
    {
        this.aSolidFill = openXmlCompositeElement.GetFirstChild<A.SolidFill>();
        if (this.aSolidFill == null)
        {
            this.aGradFill = openXmlCompositeElement!.GetFirstChild<A.GradientFill>();
            if (this.aGradFill == null)
            {
                this.InitPictureFillOr();
            }
        }
    }

    private bool HasSolidFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.SolidFill>() != null;
    }

    private bool HasGradientFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.GradientFill>() != null;
    }

    private bool HasBlipFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.BlipFill>() != null;
    }

    private bool HasPatternFill()
    {
        return openXmlCompositeElement.GetFirstChild<A.PatternFill>() != null;
    }

    private FillType GetFillType()
    {
        if (this.HasSolidFill())
        {
            return FillType.Solid;
        }

        if (this.HasGradientFill())
        {
            return FillType.Gradient;
        }

        if (this.HasBlipFill())
        {
            return FillType.Picture;
        }

        if (this.HasPatternFill())
        {
            return FillType.Pattern;
        }

        if (openXmlCompositeElement.Ancestors<P.Shape>().FirstOrDefault()?.UseBackgroundFill is not null)
        {
            return FillType.SlideBackground;
        }

        return FillType.NoFill;
    }

    private string? ColorHexOrNullOf(string schemeColor)
    {
        var openXmlPart = openXmlCompositeElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aColorScheme = GetColorScheme(openXmlPart);

        var aColor2Type = aColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        return aColor2Type?.RgbColorModelHex?.Val?.Value
               ?? aColor2Type?.SystemColor?.LastColor?.Value;
    }

    private void InitPictureFillOr()
    {
        this.aBlipFill = openXmlCompositeElement.GetFirstChild<A.BlipFill>();

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