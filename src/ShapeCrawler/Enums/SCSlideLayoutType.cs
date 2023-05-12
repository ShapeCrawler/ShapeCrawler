using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using System;

namespace ShapeCrawler.Enums;

/// <summary>
/// Enumerate all posible layout types.
/// </summary>
public class SCSlideLayoutType : Enumeration<SCSlideLayoutType>, ISlideLayoutType
{
    /// <summary>
    /// Chart.
    /// </summary>
    public static readonly SCSlideLayoutType Chart = new("chart", nameof(Chart));

    /// <summary>
    /// Clip Art and Text.
    /// </summary>
    public static readonly SCSlideLayoutType ClipArtAndText = new("clipArtAndTx", nameof(ClipArtAndText));

    /// <summary>
    /// Clip Art and Vertical Text.
    /// </summary>
    public static readonly SCSlideLayoutType ClipArtAndVerticalText = new("clipArtAndVertTx", nameof(ClipArtAndVerticalText));

    /// <summary>
    /// Four Objects.
    /// </summary>
    public static readonly SCSlideLayoutType FourObjects = new("fourObj", nameof(FourObjects));

    /// <summary>
    /// Object and Two Object.
    /// </summary>
    public static readonly SCSlideLayoutType ObjectAndTwoObject = new("objAndTwoObj", nameof(ObjectAndTwoObject));

    /// <summary>
    /// Object.
    /// </summary>
    public static readonly SCSlideLayoutType Object = new("objOnly", nameof(Object));

    /// <summary>
    /// Picture and Caption.
    /// </summary>
    public static readonly SCSlideLayoutType PictureAndCaption = new("picTx", nameof(PictureAndCaption));

    /// <summary>
    /// Section Header.
    /// </summary>
    public static readonly SCSlideLayoutType SectionHeader = new("secHead", nameof(SectionHeader));

    /// <summary>
    /// Slide Layout Type Enumeration (Chart and Text).
    /// </summary>
    public static readonly SCSlideLayoutType ChartAndText = new("chartAndTx", nameof(ChartAndText));

    /// <summary>
    /// Slide Layout Type Enumeration (Blank).
    /// </summary>
    public static readonly SCSlideLayoutType Blank = new("blank", nameof(Blank));

    /// <summary>
    /// Slide Layout Type Enumeration (Custom).
    /// </summary>
    public static readonly SCSlideLayoutType Custom = new("cust", nameof(Custom));

    /// <summary>
    /// Slide Layout Type Enumeration (Diagram).
    /// </summary>
    public static readonly SCSlideLayoutType Diagram = new("dgm", nameof(Diagram));

    /// <summary>
    /// Slide Layout Type Enumeration (Media and Text).
    /// </summary>
    public static readonly SCSlideLayoutType MediaAndText = new("mediaAndTx", nameof(MediaAndText));

    /// <summary>
    /// Slide Layout Type Enumeration (Object and Text).
    /// </summary>
    public static readonly SCSlideLayoutType ObjectAndText = new("objAndTx", nameof(ObjectAndText));

    /// <summary>
    /// Slide Layout Type Enumeration (Object over Text).
    /// </summary>
    public static readonly SCSlideLayoutType ObjectOverText = new("objOverTx", nameof(ObjectOverText));

    /// <summary>
    /// Slide Layout Type Enumeration (Table).
    /// </summary>
    public static readonly SCSlideLayoutType Table = new("tbl", nameof(Table));

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Chart).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndChart = new("tx", nameof(TextAndChart));

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Media).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndMedia = new("txAndClipArt", nameof(TextAndMedia));

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Object).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndObject = new("txAndMedia", nameof(TextAndObject));

    /// <summary>
    /// Slide Layout Type Enumeration (Text over Object).
    /// </summary>
    public static readonly SCSlideLayoutType TextOverObject = new("txAndTwoObj", nameof(TextOverObject));

    /// <summary>
    /// Slide Layout Type Enumeration (Text).
    /// </summary>
    public static readonly SCSlideLayoutType Text = new("twoTxTwoObj", nameof(Text));

    /// <summary>
    /// Slide Layout Type Enumeration (Title Only).
    /// </summary>
    public static readonly SCSlideLayoutType TitleOnly = new("titleOnly", nameof(TitleOnly));

    /// <summary>
    /// Slide Layout Type Enumeration (Title).
    /// </summary>
    public static readonly SCSlideLayoutType Title = new("title", nameof(Title));

    /// <summary>
    /// Slide Layout Type Enumeration (Two Column Text).
    /// </summary>
    public static readonly SCSlideLayoutType TwoColumnText = new("twoColTx", nameof(TwoColumnText));

    /// <summary>
    /// Text and Clip Art.
    /// </summary>
    public static readonly SCSlideLayoutType TextAndClipArt = new("txAndChart", nameof(TextAndClipArt));

    /// <summary>
    /// Text and Two Objects.
    /// </summary>
    public static readonly SCSlideLayoutType TextAndTwoObjects = new("txAndObj", nameof(TextAndTwoObjects));

    /// <summary>
    /// Title and Object.
    /// </summary>
    public static readonly SCSlideLayoutType TitleAndObject = new("obj", nameof(TitleAndObject));

    /// <summary>
    /// Title, Object, and Caption.
    /// </summary>
    public static readonly SCSlideLayoutType TitleObjectAndCaption = new("objTx", nameof(TitleObjectAndCaption));

    /// <summary>
    /// Two Objects and Object.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsAndObject = new("twoObjAndObj", nameof(TwoObjectsAndObject));

    /// <summary>
    /// Two Objects and Text.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsAndText = new("twoObjAndTx", nameof(TwoObjectsAndText));

    /// <summary>
    /// Two Objects over Text.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsOverText = new("twoObjOverTx", nameof(TwoObjectsOverText));

    /// <summary>
    /// Two Objects.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjects = new("twoObj", nameof(TwoObjects));

    /// <summary>
    /// Vertical Text.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalText = new("vertTx", nameof(VerticalText));

    /// <summary>
    /// Vertical Title and Text Over Chart.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalTitleAndTextOverChart = new("vertTitleAndTx", nameof(VerticalTitleAndTextOverChart));

    /// <summary>
    /// Vertical Title and Text.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalTitleAndText = new("txOverObj", nameof(VerticalTitleAndText));

    internal SCSlideLayoutType(string value, string name) : base(value, name)
    {
    }
}
