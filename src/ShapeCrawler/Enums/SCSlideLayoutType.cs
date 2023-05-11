namespace ShapeCrawler.Enums;

/// <summary>
/// Slide layout type.
/// </summary>
public interface ISlideLayoutType
{
    /// <summary>
    /// Gets the layout type.
    /// </summary>
    string Type { get; }
}

/// <summary>
/// Enumerate all posible layout types.
/// </summary>
public class SCSlideLayoutType : ISlideLayoutType
{
    /// <summary>
    /// Chart.
    /// </summary>
    public static readonly SCSlideLayoutType Chart = new("chart");

    /// <summary>
    /// Clip Art and Text.
    /// </summary>
    public static readonly SCSlideLayoutType ClipArtAndText = new("clipArtAndTx");

    /// <summary>
    /// Clip Art and Vertical Text.
    /// </summary>
    public static readonly SCSlideLayoutType ClipArtAndVerticalText = new("clipArtAndVertTx");

    /// <summary>
    /// Four Objects.
    /// </summary>
    public static readonly SCSlideLayoutType FourObjects = new("fourObj");

    /// <summary>
    /// Object and Two Object.
    /// </summary>
    public static readonly SCSlideLayoutType ObjectAndTwoObject = new("objAndTwoObj");

    /// <summary>
    /// Object.
    /// </summary>
    public static readonly SCSlideLayoutType Object = new("objOnly");

    /// <summary>
    /// Picture and Caption.
    /// </summary>
    public static readonly SCSlideLayoutType PictureAndCaption = new("picTx");

    /// <summary>
    /// Section Header.
    /// </summary>
    public static readonly SCSlideLayoutType SectionHeader = new("secHead");

    /// <summary>
    /// Slide Layout Type Enumeration (Chart and Text).
    /// </summary>
    public static readonly SCSlideLayoutType ChartAndText = new("chartAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Blank).
    /// </summary>
    public static readonly SCSlideLayoutType Blank = new("blank");

    /// <summary>
    /// Slide Layout Type Enumeration (Custom).
    /// </summary>
    public static readonly SCSlideLayoutType Custom = new("cust");

    /// <summary>
    /// Slide Layout Type Enumeration (Diagram).
    /// </summary>
    public static readonly SCSlideLayoutType Diagram = new("dgm");

    /// <summary>
    /// Slide Layout Type Enumeration (Media and Text).
    /// </summary>
    public static readonly SCSlideLayoutType MediaAndText = new("mediaAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Object and Text).
    /// </summary>
    public static readonly SCSlideLayoutType ObjectAndText = new("objAndTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Object over Text).
    /// </summary>
    public static readonly SCSlideLayoutType ObjectOverText = new("objOverTx");

    /// <summary>
    /// Slide Layout Type Enumeration (Table).
    /// </summary>
    public static readonly SCSlideLayoutType Table = new("tbl");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Chart).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndChart = new("tx");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Media).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndMedia = new("txAndClipArt");

    /// <summary>
    /// Slide Layout Type Enumeration (Text and Object).
    /// </summary>
    public static readonly SCSlideLayoutType TextAndObject = new("txAndMedia");

    /// <summary>
    /// Slide Layout Type Enumeration (Text over Object).
    /// </summary>
    public static readonly SCSlideLayoutType TextOverObject = new("txAndTwoObj");

    /// <summary>
    /// Slide Layout Type Enumeration (Text).
    /// </summary>
    public static readonly SCSlideLayoutType Text = new("twoTxTwoObj");

    /// <summary>
    /// Slide Layout Type Enumeration (Title Only).
    /// </summary>
    public static readonly SCSlideLayoutType TitleOnly = new("titleOnly");

    /// <summary>
    /// Slide Layout Type Enumeration (Title).
    /// </summary>
    public static readonly SCSlideLayoutType Title = new("title");

    /// <summary>
    /// Slide Layout Type Enumeration (Two Column Text).
    /// </summary>
    public static readonly SCSlideLayoutType TwoColumnText = new("twoColTx");

    /// <summary>
    /// Text and Clip Art.
    /// </summary>
    public static readonly SCSlideLayoutType TextAndClipArt = new("txAndChart");

    /// <summary>
    /// Text and Two Objects.
    /// </summary>
    public static readonly SCSlideLayoutType TextAndTwoObjects = new("txAndObj");

    /// <summary>
    /// Title and Object.
    /// </summary>
    public static readonly SCSlideLayoutType TitleAndObject = new("obj");

    /// <summary>
    /// Title, Object, and Caption.
    /// </summary>
    public static readonly SCSlideLayoutType TitleObjectAndCaption = new("objTx");

    /// <summary>
    /// Two Objects and Object.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsAndObject = new("twoObjAndObj");

    /// <summary>
    /// Two Objects and Text.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsAndText = new("twoObjAndTx");

    /// <summary>
    /// Two Objects over Text.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjectsOverText = new("twoObjOverTx");

    /// <summary>
    /// Two Objects.
    /// </summary>
    public static readonly SCSlideLayoutType TwoObjects = new("twoObj");

    /// <summary>
    /// Vertical Text.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalText = new("vertTx");

    /// <summary>
    /// Vertical Title and Text Over Chart.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalTitleAndTextOverChart = new("vertTitleAndTx");

    /// <summary>
    /// Vertical Title and Text.
    /// </summary>
    public static readonly SCSlideLayoutType VerticalTitleAndText = new("txOverObj");

    internal SCSlideLayoutType(string type)
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
