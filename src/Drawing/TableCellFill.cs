using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal class TableCellFill : IShapeFill
{
    private readonly A.TableCellProperties aTableCellProperties;
    private FillType fillType;
    private bool isDirty;
    private string? hexSolidColor;
    private ShapeFillImage? pictureImage;
    private A.SolidFill? sdkASolidFill;
    private A.GradientFill? sdkAGradFill;
    private A.PatternFill? sdkAPattFill;
    private A.BlipFill? aBlipFill;

    internal TableCellFill(A.TableCellProperties aTableCellProperties)
    {
        this.aTableCellProperties = aTableCellProperties;
        this.isDirty = true;
    }

    public string? Color => this.GetHexSolidColor();
    
    public double Alpha { get; }
    
    public double LuminanceModulation { get; }
    
    public double LuminanceOffset { get; }

    public IImage? Picture => this.GetPicture();

    public FillType Type => this.GetFillType();

    public void SetPicture(Stream image)
    {
        var openXmlPart = this.aTableCellProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (this.isDirty)
        {
            this.Initialize();
        }

        if (this.Type == FillType.Picture)
        {
            this.pictureImage!.Update(image);
        }
        else
        {
            (var rId, _) = openXmlPart.AddImagePart(image, "image/png");

            // This could be refactored to DRY vs SlideShapes.CreatePPicture.
            // In the process, the image could be de-duped also.
            this.aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            this.aBlipFill.Append(new A.Blip { Embed = rId });
            this.aBlipFill.Append(aStretch);

            this.aTableCellProperties.Append(this.aBlipFill);

            this.sdkASolidFill?.Remove();
            this.aBlipFill = null;
            this.sdkAGradFill?.Remove();
            this.sdkAGradFill = null;
            this.sdkAPattFill?.Remove();
            this.sdkAPattFill = null;
        }

        this.isDirty = true;
    }

    public void SetColor(string hex)
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        this.aTableCellProperties.AddSolidFill(hex);
        
        this.isDirty = true;
    }


    public void SetNoFill()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        this.aTableCellProperties.AddNoFill();

        this.isDirty = true;
    }

    private void InitSlideBackgroundFillOr()
    {
        this.fillType = FillType.NoFill;
    }
    
    private FillType GetFillType()
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
        this.sdkASolidFill = this.aTableCellProperties.GetFirstChild<A.SolidFill>();
        if (this.sdkASolidFill != null)
        {
            var aRgbColorModelHex = this.sdkASolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                this.hexSolidColor = hexColor;
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
        this.sdkAGradFill = this.aTableCellProperties!.GetFirstChild<A.GradientFill>();
        if (this.sdkAGradFill != null)
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
        var openXmlPart = this.aTableCellProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        this.aBlipFill = this.aTableCellProperties.GetFirstChild<A.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var blipEmbedValue = this.aBlipFill.Blip?.Embed?.Value;
            if (blipEmbedValue != null)
            {
                var imagePart = (ImagePart)openXmlPart.GetPartById(blipEmbedValue);
                var image = new ShapeFillImage(this.aBlipFill.Blip!, imagePart);
                this.pictureImage = image;
                this.fillType = FillType.Picture;
            }
        }
        else
        {
            this.InitPatternFillOr();
        }
    }

    private void InitPatternFillOr()
    {
        this.sdkAPattFill = this.aTableCellProperties.GetFirstChild<A.PatternFill>();
        if (this.sdkAPattFill != null)
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
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.hexSolidColor;
    }

    private ShapeFillImage? GetPicture()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.pictureImage;
    }
}