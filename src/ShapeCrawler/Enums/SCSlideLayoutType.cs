using System.Runtime.Serialization;

namespace ShapeCrawler;

/// <summary>
///     Represents the slide layout type.
/// </summary>
public enum SCSlideLayoutType
{
    /// <summary>
    ///     Custom.
    /// </summary>
    [EnumMember(Value = "cust")]
    Custom,

    /// <summary>
    ///     Title.
    /// </summary>
    Title,

    /// <summary>
    ///     Text.
    /// </summary>
    Text,

    /// <summary>
    ///     Two Column Text.
    /// </summary>
    TwoColumnText,

    /// <summary>
    ///     Table.
    /// </summary>
    Table,

    /// <summary>
    ///     Text and Chart.
    /// </summary>
    TextAndChart,

    /// <summary>
    ///     Chart and Text.
    /// </summary>
    ChartAndText,

    /// <summary>
    ///     Diagram.
    /// </summary>
    Diagram,

    /// <summary>
    ///     Chart.
    /// </summary>
    Chart,

    /// <summary>
    ///     Text and Clip Art.
    /// </summary>
    TextAndClipArt,

    /// <summary>
    ///     Clip Art and Text.
    /// </summary>
    ClipArtAndText,

    /// <summary>
    ///     Title Only.
    /// </summary>
    TitleOnly,

    /// <summary>
    ///     Blank.
    /// </summary>
    Blank,

    /// <summary>
    ///     Text and Object.
    /// </summary>
    TextAndObject,

    /// <summary>
    ///     Object and Text.
    /// </summary>
    ObjectAndText,

    /// <summary>
    ///     Object.
    /// </summary>
    Object,

    /// <summary>
    ///     Title and Object.
    /// </summary>
    TitleAndObject,

    /// <summary>
    ///     Text and Media.
    /// </summary>
    TextAndMedia,

    /// <summary>
    ///     Media and Text.
    /// </summary>
    MediaAndText,

    /// <summary>
    ///     Object over Text.
    /// </summary>
    ObjectOverText,

    /// <summary>
    ///     Text over Object.
    /// </summary>
    TextOverObject,

    /// <summary>
    ///     Text and Two Objects.
    /// </summary>
    TextAndTwoObjects,

    /// <summary>
    ///     Two Objects and Text.
    /// </summary>
    TwoObjectsAndText,

    /// <summary>
    ///     Two Objects over Text.
    /// </summary>
    TwoObjectsOverText,

    /// <summary>
    ///     Four Objects.
    /// </summary>
    FourObjects,

    /// <summary>
    ///     Vertical Text.
    /// </summary>
    VerticalText,

    /// <summary>
    ///     Clip Art and Vertical Text.
    /// </summary>
    ClipArtAndVerticalText,

    /// <summary>
    ///     Vertical Title and Text.
    /// </summary>
    VerticalTitleAndText,

    /// <summary>
    ///     Vertical Title and Text Over Chart.
    /// </summary>
    VerticalTitleAndTextOverChart,

    /// <summary>
    ///     Two Objects.
    /// </summary>
    TwoObjects,

    /// <summary>
    ///     Object and Two Object.
    /// </summary>
    ObjectAndTwoObject,

    /// <summary>
    ///     Two Objects and Object.
    /// </summary>
    TwoObjectsAndObject,

    /// <summary>
    ///     Section Header.
    /// </summary>
    SectionHeader,

    /// <summary>
    ///     Title, Object, and Caption.
    /// </summary>
    TitleObjectAndCaption,

    /// <summary>
    ///     Picture and Caption.
    /// </summary>
    PictureAndCaption
}