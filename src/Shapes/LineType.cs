#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     PowerPoint line and connector type.
/// </summary>
public enum LineType
{
    /// <summary>Standard line without arrowheads.</summary>
    Line,

    /// <summary>Standard line with an arrowhead at the end.</summary>
    Arrow,

    /// <summary>Standard line with arrowheads at both ends.</summary>
    DoubleArrow,

    /// <summary>Elbow connector.</summary>
    ElbowConnector,

    /// <summary>Elbow connector with an arrowhead at the end.</summary>
    ElbowArrowConnector,

    /// <summary>Elbow connector with arrowheads at both ends.</summary>
    ElbowDoubleArrowConnector,

    /// <summary>Curved connector.</summary>
    CurvedConnector,

    /// <summary>Curved connector with an arrowhead at the end.</summary>
    CurvedArrowConnector,

    /// <summary>Curved connector with arrowheads at both ends.</summary>
    CurvedDoubleArrowConnector,

    /// <summary>Curve drawn with the Curve tool.</summary>
    Curve,

    /// <summary>Open freeform line drawn with the Freeform Shape tool.</summary>
    FreeformShape,

    /// <summary>Freehand line drawn with the Scribble tool.</summary>
    Scribble,
}