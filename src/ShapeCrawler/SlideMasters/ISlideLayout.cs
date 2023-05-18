using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets parent Slide Master.
    /// </summary>
    ISlideMaster SlideMaster { get; }

    /// <summary>
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets layout type.
    /// </summary>
    SCSlideLayoutType Type { get; }

    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }
}

internal sealed class SCSlideLayout : SlideStructure, ISlideLayout
{
    private static readonly Dictionary<string, SCSlideLayoutType> TypeMapping = new()
    {
        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_SlideLayoutType_topic_ID0EKTIIB.html
        { "blank", SCSlideLayoutType.Blank },
        { "chart", SCSlideLayoutType.Chart },
        { "chartAndTx", SCSlideLayoutType.ChartAndText },
        { "clipArtAndTx", SCSlideLayoutType.ClipArtAndText },
        { "clipArtAndVertTx", SCSlideLayoutType.ClipArtAndVerticalText },
        { "cust", SCSlideLayoutType.Custom },
        { "dgm", SCSlideLayoutType.Diagram },
        { "fourObj", SCSlideLayoutType.FourObjects },
        { "mediaAndTx", SCSlideLayoutType.MediaAndText },
        { "obj", SCSlideLayoutType.Object },
        { "objAndTwoObj", SCSlideLayoutType.ObjectAndTwoObjects },
        { "objAndTx", SCSlideLayoutType.ObjectAndText },
        { "objOnly", SCSlideLayoutType.ObjectOnly },
        { "objOverTx", SCSlideLayoutType.ObjectOverText },
        { "objTx", SCSlideLayoutType.ObjectText },
        { "picTx", SCSlideLayoutType.PictureAndCaption },
        { "secHead", SCSlideLayoutType.SectionHeader },
        { "tbl", SCSlideLayoutType.Table },
        { "title", SCSlideLayoutType.Title },
        { "titleOnly", SCSlideLayoutType.TitleOnly },
        { "twoColTx", SCSlideLayoutType.TwoColumnText },
        { "twoObj", SCSlideLayoutType.TwoObjects },
        { "twoObjAndObj", SCSlideLayoutType.TwoObjectsAndObject },
        { "twoObjAndTx", SCSlideLayoutType.TwoObjectsAndText },
        { "twoObjOverTx", SCSlideLayoutType.TwoObjectsOverText },
        { "twoTxTwoObj", SCSlideLayoutType.TwoTextAndTwoObjects },
        { "tx", SCSlideLayoutType.Text },
        { "txAndChart", SCSlideLayoutType.TextAndChart },
        { "txAndClipArt", SCSlideLayoutType.TextAndClipArt },
        { "txAndMedia", SCSlideLayoutType.TextAndMedia },
        { "txAndObj", SCSlideLayoutType.TextAndObject },
        { "txAndTwoObj", SCSlideLayoutType.TextAndTwoObjects },
        { "txOverObj", SCSlideLayoutType.TextOverObject },
        { "vertTitleAndTx", SCSlideLayoutType.VerticalTitleAndText },
        { "vertTitleAndTxOverChart", SCSlideLayoutType.VerticalTitleAndTextOverChart },
        { "vertTx", SCSlideLayoutType.VerticalText }
    };

    private readonly ResettableLazy<ShapeCollection> shapes;
    private readonly SCSlideMaster slideMaster;

    internal SCSlideLayout(SCSlideMaster slideMaster, SlideLayoutPart slideLayoutPart, int number)
        : base(slideMaster.Presentation)
    {
        this.slideMaster = slideMaster;
        this.SlideLayoutPart = slideLayoutPart;
        this.shapes = new ResettableLazy<ShapeCollection>(() =>
            new ShapeCollection(slideLayoutPart, this));
        this.Number = number;
    }

    public string Name => this.GetName();

    public SCSlideLayoutType Type => this.GetLayoutType();

    public ISlideMaster SlideMaster => this.slideMaster;

    public override int Number { get; set; }

    public override IShapeCollection Shapes => this.shapes.Value;

    internal SlideLayoutPart SlideLayoutPart { get; }

    internal SCSlideMaster SlideMasterInternal => (SCSlideMaster)this.SlideMaster;

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.SlideLayoutPart;

    private string GetName()
    {
        return this.SlideLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    }

    private SCSlideLayoutType GetLayoutType()
    {
        return TypeMapping[this.SlideLayoutPart.SlideLayout.Type!];
    }
}