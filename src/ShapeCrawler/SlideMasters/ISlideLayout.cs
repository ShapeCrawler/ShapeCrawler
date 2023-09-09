using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;

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
    SCSlideLayoutType Type { get; }

    /// <summary>
    ///     Gets layout name.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets layout shape collection.
    /// </summary>
    IReadOnlyShapes Shapes { get; }

    ISlideMaster SlideMaster { get; }
}

internal sealed class SlideLayout : ISlideLayout
{
    private readonly SlideLayoutPart sdkLayoutPart;

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

    internal SlideLayout(SlideLayoutPart sdkLayoutPart)
        : this(sdkLayoutPart, new SlideMaster(sdkLayoutPart.SlideMasterPart!))
    {
    }

    private SlideLayout(SlideLayoutPart sdkLayoutPart, ISlideMaster slideMaster)
    {
        this.sdkLayoutPart = sdkLayoutPart;
        this.SlideMaster = slideMaster;
        this.Shapes = new LayoutShapes(this.sdkLayoutPart);
    }

    public string Name => this.GetName();

    public IReadOnlyShapes Shapes { get; }

    public ISlideMaster SlideMaster { get; }

    public SCSlideLayoutType Type => this.GetLayoutType();

    private string GetName()
    {
        return this.sdkLayoutPart.SlideLayout.CommonSlideData!.Name!.Value!;
    }

    private SCSlideLayoutType GetLayoutType()
    {
        return TypeMapping[this.sdkLayoutPart.SlideLayout.Type!];
    }

    internal SlideLayoutPart SDKSlideLayoutPart()
    {
        return this.sdkLayoutPart;
    }
}