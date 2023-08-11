using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.AutoShapes;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

internal class SCShapeFill : IShapeFill
{
    private readonly TypedOpenXmlCompositeElement properties;
    private BooleanValue? useBgFill;
    private SCFillType fillType;
    private bool isDirty;
    private string? hexSolidColor;
    private SCImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private A.BlipFill? aBlipFill;

    internal SCShapeFill(TypedOpenXmlCompositeElement properties)
    {
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
            this.pictureImage!.UpdateImage(imageStream);
        }
        else
        {
            var rId = this.slideTypedOpenXmlPart.AddImagePart(imageStream);

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
        this.aBlipFill = this.properties.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var image = SCImage.ForAutoShapeFill(this.slideTypedOpenXmlPart, this.aBlipFill, this.imageParts);
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