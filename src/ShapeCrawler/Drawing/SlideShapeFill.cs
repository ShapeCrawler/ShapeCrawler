using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal record SlideShapeFill : IShapeFill
{
    private readonly TypedOpenXmlCompositeElement properties;
    private BooleanValue? useBgFill;
    private FillType fillType;
    private string? hexSolidColor;
    private SlidePictureImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private A.BlipFill? aBlipFill;
    private readonly SlidePart sdkSlidePart;

    internal SlideShapeFill(SlidePart sdkSlidePart, TypedOpenXmlCompositeElement properties, BooleanValue? useBgFill)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.properties = properties;
        this.useBgFill = useBgFill;
    }

    public string? Color
    {
        get
        {
            this.aSolidFill = this.properties.GetFirstChild<A.SolidFill>();
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

    private string? ColorHexOrNullOf(string schemeColor)
    {
        var aColorScheme = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!.ColorScheme!;
        var aColor2Type = aColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == schemeColor);
        var hex = aColor2Type?.RgbColorModelHex?.Val?.Value ?? aColor2Type?.SystemColor?.LastColor?.Value;
        
        if (hex != null)
        {
            return hex;
        }

        if (hex == null)
        {
            // GetThemeMappedColor
            var pColorMap = this.sdkSlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster.ColorMap;
            var targetSchemeColor = pColorMap?.GetAttributes().FirstOrDefault(a => a.LocalName == schemeColor)!;

            var attrValue = targetSchemeColor!.Value;
            aColor2Type = aColorScheme.Elements<A.Color2Type>().FirstOrDefault(c => c.LocalName == attrValue.Value);
            return aColor2Type?.RgbColorModelHex?.Val?.Value ?? aColor2Type?.SystemColor?.LastColor?.Value;
        }

        return null;
    }

    public double AlphaPercentage { get; }
    public double LuminanceModulation { get; }
    public double LuminanceOffset { get; }

    public IImage? Picture => this.GetPicture();

    public FillType Type => this.GetFillType();

    public void SetPicture(Stream image)
    {
        this.Initialize();

        if (this.Type == FillType.Picture)
        {
            this.pictureImage!.Update(image);
        }
        else
        {
            var rId = sdkSlidePart.AddImagePart(image);

            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
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
    }

    public void SetColor(string hex)
    {
        this.Initialize();
        this.properties.AddASolidFill(hex);
        this.useBgFill = false;
    }

    private void InitSlideBackgroundFillOr()
    {
        if (this.useBgFill is not null && this.useBgFill)
        {
            this.fillType = FillType.SlideBackground;
        }
        else
        {
            this.fillType = FillType.NoFill;
        }
    }

    private FillType GetFillType()
    {
        this.Initialize();
        return this.fillType;
    }

    private void Initialize()
    {
        this.InitSolidFillOr();
    }

    private void InitSolidFillOr()
    {
        this.aSolidFill = this.properties.GetFirstChild<A.SolidFill>();
        if (this.aSolidFill != null)
        {
            var aRgbColorModelHex = this.aSolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                this.hexSolidColor = hexColor;
            }
            else
            {
                // TODO: get hex color from scheme
                var schemeColor = this.aSolidFill.SchemeColor;
            }

            this.fillType = FillType.Solid;
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
            this.fillType = FillType.Gradient;
        }
        else
        {
            this.InitPictureFillOr();
        }
    }

    private void InitPictureFillOr()
    {
        this.aBlipFill = this.properties.GetFirstChild<A.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var image = new SlidePictureImage(this.sdkSlidePart, this.aBlipFill.Blip!);
            this.pictureImage = image;
            this.fillType = FillType.Picture;
        }
        else
        {
            this.InitPatternFillOr();
        }
    }

    private void InitPatternFillOr()
    {
        this.aPattFill = this.properties.GetFirstChild<A.PatternFill>();
        if (this.aPattFill != null)
        {
            this.fillType = FillType.Pattern;
        }
        else
        {
            this.InitSlideBackgroundFillOr();
        }
    }

    private string? GetHexSolidColor()
    {
        this.Initialize();

        return this.hexSolidColor;
    }

    private SlidePictureImage? GetPicture()
    {
        this.Initialize();

        return this.pictureImage;
    }
}