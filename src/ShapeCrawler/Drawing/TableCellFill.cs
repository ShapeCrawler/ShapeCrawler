using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class TableCellFill : IShapeFill
{
    private readonly SlidePart sdkSlidePart;
    private readonly TypedOpenXmlCompositeElement cellProperties;
    private readonly TableCell parentTableCell;
    private BooleanValue? useBgFill;
    private SCFillType fillType;
    private bool isDirty;
    private string? hexSolidColor;
    private CellFillImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private A.BlipFill? aBlipFill;

    internal TableCellFill(SlidePart sdkSlidePart, A.TableCellProperties cellProperties, TableCell parentTableCell)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.cellProperties = cellProperties;
        this.parentTableCell = parentTableCell;
        this.isDirty = true;
    }

    public string? Color => this.GetHexSolidColor();
    
    public double AlphaPercentage { get; }
    public double LuminanceModulation { get; }
    public double LuminanceOffset { get; }

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
            this.pictureImage!.Update(imageStream);
        }
        else
        {
            var rId = this.sdkSlidePart.AddImagePart(imageStream);

            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.cellProperties.Append(aBlipFill);

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

        this.cellProperties.AddASolidFill(hex);
        
        this.useBgFill = false;

        this.isDirty = true;
    }

    private void InitSlideBackgroundFillOr()
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
        this.aSolidFill = this.cellProperties.GetFirstChild<A.SolidFill>();
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
        this.aGradFill = this.cellProperties!.GetFirstChild<A.GradientFill>();
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
        this.aBlipFill = this.cellProperties.GetFirstChild<A.BlipFill>();

        if (this.aBlipFill is not null)
        {
            var blipEmbedValue = aBlipFill.Blip?.Embed?.Value;
            if (blipEmbedValue != null)
            {
                var imagePart = (ImagePart)this.sdkSlidePart.GetPartById(blipEmbedValue);
                var image = new CellFillImage(this.aBlipFill, imagePart, this);
                this.pictureImage = image;
                this.fillType = SCFillType.Picture;
            }
        }
        else
        {
            this.InitPatternFillOr();
        }
    }

    private void InitPatternFillOr()
    {
        this.aPattFill = this.cellProperties.GetFirstChild<A.PatternFill>();
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

    private CellFillImage? GetPicture()
    {
        if (this.isDirty)
        {
            this.Initialize();
        }

        return this.pictureImage;
    }
}