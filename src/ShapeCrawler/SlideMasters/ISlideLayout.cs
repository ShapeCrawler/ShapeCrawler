using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a Slide Layout.
/// </summary>
public interface ISlideLayout
{
    /// <summary>
    ///     Gets layout type.
    /// </summary>
    SlideLayoutType Type { get; }

    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets layout shape collection.
    /// </summary>
    IShapes Shapes { get; }

    ISlideMaster SlideMaster { get; }
}

internal sealed class SlideLayout : ISlideLayout
{
    private readonly SlideLayoutPart sdkLayoutPart;
    private static readonly Dictionary<string, SlideLayoutType> TypeMapping = new()
    {
        // https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_SlideLayoutType_topic_ID0EKTIIB.html
        { "blank", SlideLayoutType.Blank },
        { "chart", SlideLayoutType.Chart },
        { "chartAndTx", SlideLayoutType.ChartAndText },
        { "clipArtAndTx", SlideLayoutType.ClipArtAndText },
        { "clipArtAndVertTx", SlideLayoutType.ClipArtAndVerticalText },
        { "cust", SlideLayoutType.Custom },
        { "dgm", SlideLayoutType.Diagram },
        { "fourObj", SlideLayoutType.FourObjects },
        { "mediaAndTx", SlideLayoutType.MediaAndText },
        { "obj", SlideLayoutType.Object },
        { "objAndTwoObj", SlideLayoutType.ObjectAndTwoObjects },
        { "objAndTx", SlideLayoutType.ObjectAndText },
        { "objOnly", SlideLayoutType.ObjectOnly },
        { "objOverTx", SlideLayoutType.ObjectOverText },
        { "objTx", SlideLayoutType.ObjectText },
        { "picTx", SlideLayoutType.PictureAndCaption },
        { "secHead", SlideLayoutType.SectionHeader },
        { "tbl", SlideLayoutType.Table },
        { "title", SlideLayoutType.Title },
        { "titleOnly", SlideLayoutType.TitleOnly },
        { "twoColTx", SlideLayoutType.TwoColumnText },
        { "twoObj", SlideLayoutType.TwoObjects },
        { "twoObjAndObj", SlideLayoutType.TwoObjectsAndObject },
        { "twoObjAndTx", SlideLayoutType.TwoObjectsAndText },
        { "twoObjOverTx", SlideLayoutType.TwoObjectsOverText },
        { "twoTxTwoObj", SlideLayoutType.TwoTextAndTwoObjects },
        { "tx", SlideLayoutType.Text },
        { "txAndChart", SlideLayoutType.TextAndChart },
        { "txAndClipArt", SlideLayoutType.TextAndClipArt },
        { "txAndMedia", SlideLayoutType.TextAndMedia },
        { "txAndObj", SlideLayoutType.TextAndObject },
        { "txAndTwoObj", SlideLayoutType.TextAndTwoObjects },
        { "txOverObj", SlideLayoutType.TextOverObject },
        { "vertTitleAndTx", SlideLayoutType.VerticalTitleAndText },
        { "vertTitleAndTxOverChart", SlideLayoutType.VerticalTitleAndTextOverChart },
        { "vertTx", SlideLayoutType.VerticalText }
    };

    internal SlideLayout(SlideLayoutPart sdkLayoutPart)
        : this(sdkLayoutPart, new SlideMaster(sdkLayoutPart.SlideMasterPart!))
    {
    }

    private SlideLayout(SlideLayoutPart sdkLayoutPart, ISlideMaster slideMaster)
    {
        this.sdkLayoutPart = sdkLayoutPart;
        this.SlideMaster = slideMaster;
        this.Shapes = new ShapeCollection.Shapes(this.sdkLayoutPart);
    }

    public string Name => this.sdkLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    public IShapes Shapes { get; }
    public ISlideMaster SlideMaster { get; }
    public SlideLayoutType Type => TypeMapping[this.sdkLayoutPart.SlideLayout.Type!];
    internal SlideLayoutPart SDKSlideLayoutPart() => this.sdkLayoutPart;
}