namespace ShapeCrawler.Enums;

/// <summary>
/// Enumerate all posible layout types.
/// </summary>
public class SlideLayoutType
{
    /// <summary>
    /// Chart.
    /// </summary>
    public static readonly SlideLayoutType Chart = new("chart");

    /// <summary>
    /// Clip Art and Text.
    /// </summary>
    public static readonly SlideLayoutType ClipArtAndText = new("clipArtAndTx");

    /// <summary>
    /// Clip Art and Vertical Text.
    /// </summary>
    public static readonly SlideLayoutType ClipArtAndVerticalText = new("clipArtAndVertTx");

    /// <summary>
    /// Four Objects.
    /// </summary>
    public static readonly SlideLayoutType FourObjects = new("fourObj");

    /// <summary>
    /// Object and Two Object.
    /// </summary>
    public static readonly SlideLayoutType ObjectAndTwoObject = new("objAndTwoObj");

    /// <summary>
    /// Object.
    /// </summary>
    public static readonly SlideLayoutType Object = new("objOnly");

    /// <summary>
    /// Picture and Caption.
    /// </summary>
    public static readonly SlideLayoutType PictureAndCaption = new("picTx");

    /// <summary>
    /// Section Header.
    /// </summary>
    public static readonly SlideLayoutType SectionHeader = new("secHead");

    /// <summary>
    /// Slide Layout Type Enumeration (Chart and Text).
    /// </summary>
    public static readonly SlideLayoutType ChartAndText = new("chartAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Blank).
    /// </summary>
    public static readonly SlideLayoutType Blank = new("blank");

    /// <summary>
    /// Slide Layout Type Enumeration (Custom).
    /// </summary>
    public static readonly SlideLayoutType Custom = new("cust");

    /// <summary>
    /// Slide Layout Type Enumeration (Diagram).
    /// </summary>
    public static readonly SlideLayoutType Diagram = new("dgm");

    /// <summary>
    /// Slide Layout Type Enumeration (Media and Text).
    /// </summary>
    public static readonly SlideLayoutType MediaAndText = new("mediaAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Object and Text).
    /// </summary>
    public static readonly SlideLayoutType ObjectAndText = new("objAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Object over Text).
    /// </summary>
    public static readonly SlideLayoutType ObjectOverText = new("objOverTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Table).
    /// </summary>
    public static readonly SlideLayoutType Table = new("tbl");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Chart).
    /// </summary>
    public static readonly SlideLayoutType TextAndChart = new("tx");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Media).
    /// </summary>
    public static readonly SlideLayoutType TextAndMedia = new("txAndClipArt");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Object).
    /// </summary>
    public static readonly SlideLayoutType TextAndObject = new("txAndMedia");

    /// <summary>
    /// Slide Layout Type Enumeration (Text over Object).
    /// </summary>
    public static readonly SlideLayoutType TextOverObject = new("txAndTwoObj");

    /// <summary>
    /// Slide Layout Type Enumeration (Text).
    /// </summary>
    public static readonly SlideLayoutType Text = new("twoTxTwoObj");

    /// <summary>
    /// Slide Layout Type Enumeration (Title Only).
    /// </summary>
    public static readonly SlideLayoutType TitleOnly = new("titleOnly");

    /// <summary>
    /// Slide Layout Type Enumeration (Title).
    /// </summary>
    public static readonly SlideLayoutType Title = new("title");

    /// <summary>
    /// Slide Layout Type Enumeration (Two Column Text).
    /// </summary>
    public static readonly SlideLayoutType TwoColumnText = new("twoColTx");

    /// <summary>
    /// Text and Clip Art.
    /// </summary>
    public static readonly SlideLayoutType TextAndClipArt = new("txAndChart");

    /// <summary>
    /// Text and Two Objects.
    /// </summary>
    public static readonly SlideLayoutType TextAndTwoObjects = new("txAndObj");

    /// <summary>
    /// Title and Object.
    /// </summary>
    public static readonly SlideLayoutType TitleAndObject = new("obj");

    /// <summary>
    /// Title, Object, and Caption.
    /// </summary>
    public static readonly SlideLayoutType TitleObjectAndCaption = new("objTx");

    /// <summary>
    /// Two Objects and Object.
    /// </summary>
    public static readonly SlideLayoutType TwoObjectsAndObject = new("twoObjAndObj");

    /// <summary>
    /// Two Objects and Text.
    /// </summary>
    public static readonly SlideLayoutType TwoObjectsAndText = new("twoObjAndTx");

    /// <summary>
    /// Two Objects over Text.
    /// </summary>
    public static readonly SlideLayoutType TwoObjectsOverText = new("twoObjOverTx");

    /// <summary>
    /// Two Objects.
    /// </summary>
    public static readonly SlideLayoutType TwoObjects = new("twoObj");

    /// <summary>
    /// Vertical Text.
    /// </summary>
    public static readonly SlideLayoutType VerticalText = new("vertTx");

    /// <summary>
    /// Vertical Title and Text Over Chart.
    /// </summary>
    public static readonly SlideLayoutType VerticalTitleAndTextOverChart = new("vertTitleAndTx");

    /// <summary>
    /// Vertical Title and Text.
    /// </summary>
    public static readonly SlideLayoutType VerticalTitleAndText = new("txOverObj");

    internal SlideLayoutType(string type)
    {
        this.Type = type;
    }

    /// <summary>
    /// Gets the layout type.
    /// </summary>
    public string Type { get; }

    /// <inheritdoc/>
    public override string ToString()
    {
        return this.Type;
    }
}
