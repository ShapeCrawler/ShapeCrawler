using System.IO;
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
    
    private readonly SlideStructure slideObject;
    private bool isDirty;
    private string? hexSolidColor;
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
            }
            else
            {
                // TODO: get hex color from scheme
                var schemeColor = this.aSolidFill.SchemeColor;
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

    private SCImage? GetPicture()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.pictureImage;
    }
}