using System.IO;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class ShapeFill : IShapeFill
{
    private readonly Shape shape;
    private bool isDirty;
    private SCFillType fillType;
    private string? hexSolidColor;
    private SCImage? pictureImage;
    private A.SolidFill? aSolidFill;
    private A.GradientFill? aGradFill;
    private A.PatternFill? aPattFill;
    private BooleanValue? useBgFill;

    internal ShapeFill(Shape shape)
    {
        this.shape = shape;
        this.isDirty = true;
    }

    public string? HexSolidColor => this.GetHexSolidColor();

    public IImage? Picture => this.GetPicture();

    public SCFillType Type => this.GetFillType();
        
    public void SetPicture(Stream imageStream)
    {
        if (this.Type == SCFillType.Picture)
        {
            this.pictureImage!.SetImage(imageStream);
        }
        else
        {
            var rId = this.shape.SlideBase.TypedOpenXmlPart.AddImagePart(imageStream);

            var aBlipFill = new A.BlipFill();
            var aStretch = new A.Stretch();
            aStretch.Append(new A.FillRectangle());
            aBlipFill.Append(new A.Blip { Embed = rId });
            aBlipFill.Append(aStretch);

            this.shape.PShapeProperties.Append(aBlipFill);
                
            this.aSolidFill?.Remove();
            this.aGradFill?.Remove();
            this.aPattFill?.Remove();
            this.useBgFill = false;
        }

        this.isDirty = true;
    }

    public void SetHexSolidColor(string hex)
    {
        if (this.isDirty)
        {
            this.Initialize();
        }
        
        var pShape = (P.Shape)this.shape.PShapeTreesChild;
        pShape.ShapeProperties!.AddASolidFill(hex);

        this.isDirty = true;
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
        var pShape = (P.Shape)this.shape.PShapeTreesChild;
        this.aSolidFill = pShape.ShapeProperties!.GetFirstChild<A.SolidFill>();
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
            this.InitGradientFillOr(pShape);
        }
    }
        
    private void InitGradientFillOr(P.Shape pShape)
    {
        this.aGradFill = pShape.ShapeProperties!.GetFirstChild<A.GradientFill>();
        if (this.aGradFill != null)
        {
            this.fillType = SCFillType.Gradient;
        }
        else
        {
            this.InitPictureFillOr(pShape);
        }
    }

    private void InitPictureFillOr(P.Shape pShape)
    {
        var xmlPart = this.shape.SlideBase.TypedOpenXmlPart;
        var image = SCImage.ForAutoShapeFill(this.shape, xmlPart);
        if (image != null)
        {
            this.pictureImage = image;
            this.fillType = SCFillType.Picture;
        }
        else
        {
            this.InitPatternFillOr(pShape);
        }
    }

    private void InitPatternFillOr(P.Shape pShape)
    {
        this.aPattFill = this.shape.PShapeProperties.GetFirstChild<A.PatternFill>();
        if (this.aPattFill != null)
        {
            this.fillType = SCFillType.Pattern;
        }
        else
        {
            this.InitSlideBackgroundFillOr(pShape);
        }
    }

    private void InitSlideBackgroundFillOr(P.Shape pShape)
    {
        this.useBgFill = pShape.UseBackgroundFill; 
        if (this.useBgFill is not null && this.useBgFill)
        {
            this.useBgFill = pShape.UseBackgroundFill;
            this.fillType = SCFillType.SlideBackground;
        }
        else
        {
            this.fillType = SCFillType.NoFill;
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