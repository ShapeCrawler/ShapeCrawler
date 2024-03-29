using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal class TableCellFill : IShapeFill
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.TableCellProperties sdkATableCellProperties;
    private FillType fillType;
    private bool isDirty;
    private string? hexSolidColor;
    private ShapeFillImage? pictureImage;
    private A.SolidFill? sdkASolidFill;
    private A.GradientFill? sdkAGradFill;
    private A.PatternFill? sdkAPattFill;
    private A.BlipFill? sdkABlipFill;

    internal TableCellFill(OpenXmlPart sdkTypedOpenXmlPart, A.TableCellProperties sdkATableCellProperties)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.sdkATableCellProperties = sdkATableCellProperties;
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
            var rId = this.sdkTypedOpenXmlPart.AddImagePart(image);

            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.sdkATableCellProperties.Append(aBlipFill);

            this.sdkASolidFill?.Remove();
            this.sdkABlipFill = null;
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

        this.sdkATableCellProperties.AddASolidFill(hex);
        
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
        this.sdkASolidFill = this.sdkATableCellProperties.GetFirstChild<A.SolidFill>();
        if (this.sdkASolidFill != null)
        {
            var aRgbColorModelHex = this.sdkASolidFill.RgbColorModelHex;
            if (aRgbColorModelHex != null)
            {
                var hexColor = aRgbColorModelHex.Val!.ToString();
                this.hexSolidColor = hexColor;
            }
            else
            {
                // TODO: get hex color from scheme
                var schemeColor = this.sdkASolidFill.SchemeColor;
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
        this.sdkAGradFill = this.sdkATableCellProperties!.GetFirstChild<A.GradientFill>();
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
        this.sdkABlipFill = this.sdkATableCellProperties.GetFirstChild<A.BlipFill>();

        if (this.sdkABlipFill is not null)
        {
            var blipEmbedValue = this.sdkABlipFill.Blip?.Embed?.Value;
            if (blipEmbedValue != null)
            {
                var imagePart = (ImagePart)this.sdkTypedOpenXmlPart.GetPartById(blipEmbedValue);
                var image = new ShapeFillImage(this.sdkTypedOpenXmlPart, this.sdkABlipFill, imagePart);
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
        this.sdkAPattFill = this.sdkATableCellProperties.GetFirstChild<A.PatternFill>();
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