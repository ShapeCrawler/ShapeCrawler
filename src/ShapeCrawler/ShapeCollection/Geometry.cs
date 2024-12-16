#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Geometry type.
/// </summary>
public enum Geometry
{
    /// <summary>
    ///     Line.
    /// </summary>
    Line,

    /// <summary>
    ///     Line Inverse.
    /// </summary>
    LineInverse,

    /// <summary>
    ///     Triangle.
    /// </summary>
    Triangle,

    /// <summary>
    ///     Right Triangle.
    /// </summary>
    RightTriangle,

    /// <summary>
    ///     Rectangle.
    /// </summary>
    Rectangle,

    /// <summary>
    ///     Diamond.
    /// </summary>
    Diamond,

    /// <summary>
    ///     Parallelogram.
    /// </summary>
    Parallelogram,

    /// <summary>
    ///     Trapezoid.
    /// </summary>
    Trapezoid,

    /// <summary>
    ///     Non Isosceles Trapezoid.
    /// </summary>
    NonIsoscelesTrapezoid,

    /// <summary>
    ///     Pentagon.
    /// </summary>
    Pentagon,

    /// <summary>
    ///     Hexagon.
    /// </summary>
    Hexagon,

    /// <summary>
    ///     Heptagon.
    /// </summary>
    Heptagon,

    /// <summary>
    ///     Octagon.
    /// </summary>
    Octagon,

    /// <summary>
    ///     Decagon.
    /// </summary>
    Decagon,

    /// <summary>
    ///     Dodecagon.
    /// </summary>
    Dodecagon,

    /// <summary>
    ///     Star4.
    /// </summary>
    Star4,

    /// <summary>
    ///     Star5.
    /// </summary>
    Star5,

    /// <summary>
    ///     Star6.
    /// </summary>
    Star6,

    /// <summary>
    ///     Star7.
    /// </summary>
    Star7,

    /// <summary>
    ///     Star8.
    /// </summary>
    Star8,

    /// <summary>
    ///     Star10.
    /// </summary>
    Star10,

    /// <summary>
    ///     Star12.
    /// </summary>
    Star12,

    /// <summary>
    ///     Star16.
    /// </summary>
    Star16,

    /// <summary>
    ///     Star24.
    /// </summary>
    Star24,

    /// <summary>
    ///     Star32.
    /// </summary>
    Star32,

    /// <summary>
    ///     Round Rectangle.
    /// </summary>
    RoundedRectangle,

    /// <summary>
    ///     Round1Rectangle.
    /// </summary>
    SingleCornerRoundedRectangle,

    /// <summary>
    ///     Round2SameRectangle
    /// </summary>
    TopCornersRoundedRectangle,

    /// <summary>
    ///     Round2DiagonalRectangle
    /// </summary>
    DiagonalCornersRoundedRectangle,

    /// <summary>
    ///     SnipRoundRectangle
    /// </summary>
    SnipRoundRectangle,

    /// <summary>
    ///     Snip1Rectangle
    /// </summary>
    Snip1Rectangle,

    /// <summary>
    ///     Snip2SameRectangle
    /// </summary>
    Snip2SameRectangle,

    /// <summary>
    ///     Snip2DiagonalRectangle
    /// </summary>
    Snip2DiagonalRectangle,

    /// <summary>
    ///     Plaque
    /// </summary>
    Plaque,

    /// <summary>
    ///     Ellipse
    /// </summary>
    Ellipse,

    /// <summary>
    ///     Teardrop
    /// </summary>
    Teardrop,

    /// <summary>
    ///     HomePlate
    /// </summary>
    HomePlate,

    /// <summary>
    ///     Chevron
    /// </summary>
    Chevron,

    /// <summary>
    ///     PieWedge
    /// </summary>
    PieWedge,

    /// <summary>
    ///     Pie
    /// </summary>
    Pie,

    /// <summary>
    ///     BlockArc
    /// </summary>
    BlockArc,

    /// <summary>
    ///     Donut
    /// </summary>
    Donut,

    /// <summary>
    ///     NoSmoking
    /// </summary>
    NoSmoking,

    /// <summary>
    ///     RightArrow
    /// </summary>
    RightArrow,

    /// <summary>
    ///     LeftArrow
    /// </summary>
    LeftArrow,

    /// <summary>
    ///     UpArrow
    /// </summary>
    UpArrow,

    /// <summary>
    ///     DownArrow
    /// </summary>
    DownArrow,

    /// <summary>
    ///     StripedRightArrow
    /// </summary>
    StripedRightArrow,

    /// <summary>
    ///     NotchedRightArrow
    /// </summary>
    NotchedRightArrow,

    /// <summary>
    ///     BentUpArrow
    /// </summary>
    BentUpArrow,

    /// <summary>
    ///     LeftRightArrow
    /// </summary>
    LeftRightArrow,

    /// <summary>
    ///     UpDownArrow
    /// </summary>
    UpDownArrow,

    /// <summary>
    ///     LeftUpArrow
    /// </summary>
    LeftUpArrow,

    /// <summary>
    ///     LeftRightUpArrow
    /// </summary>
    LeftRightUpArrow,

    /// <summary>
    ///     QuadArrow
    /// </summary>
    QuadArrow,

    /// <summary>
    ///     LeftArrowCallout
    /// </summary>
    LeftArrowCallout,

    /// <summary>
    ///     RightArrowCallout
    /// </summary>
    RightArrowCallout,

    /// <summary>
    ///     UpArrowCallout
    /// </summary>
    UpArrowCallout,

    /// <summary>
    ///     DownArrowCallout
    /// </summary>
    DownArrowCallout,

    /// <summary>
    ///     LeftRightArrowCallout
    /// </summary>
    LeftRightArrowCallout,

    /// <summary>
    ///     UpDownArrowCallout
    /// </summary>
    UpDownArrowCallout,

    /// <summary>
    ///     QuadArrowCallout
    /// </summary>
    QuadArrowCallout,

    /// <summary>
    ///     BentArrow
    /// </summary>
    BentArrow,

    /// <summary>
    ///     UTurnArrow
    /// </summary>
    UTurnArrow,

    /// <summary>
    ///     CircularArrow
    /// </summary>
    CircularArrow,

    /// <summary>
    ///     LeftCircularArrow
    /// </summary>
    LeftCircularArrow,

    /// <summary>
    ///     LeftRightCircularArrow
    /// </summary>
    LeftRightCircularArrow,

    /// <summary>
    ///     CurvedRightArrow
    /// </summary>
    CurvedRightArrow,

    /// <summary>
    ///     CurvedLeftArrow
    /// </summary>
    CurvedLeftArrow,

    /// <summary>
    ///     CurvedUpArrow
    /// </summary>
    CurvedUpArrow,

    /// <summary>
    ///     CurvedDownArrow
    /// </summary>
    CurvedDownArrow,

    /// <summary>
    ///     SwooshArrow
    /// </summary>
    SwooshArrow,

    /// <summary>
    ///     Cube
    /// </summary>
    Cube,

    /// <summary>
    ///     Can
    /// </summary>
    Can,

    /// <summary>
    ///     LightningBolt
    /// </summary>
    LightningBolt,

    /// <summary>
    ///     Heart
    /// </summary>
    Heart,

    /// <summary>
    ///     Sun
    /// </summary>
    Sun,

    /// <summary>
    ///     Moon
    /// </summary>
    Moon,

    /// <summary>
    ///     SmileyFace
    /// </summary>
    SmileyFace,

    /// <summary>
    ///     IrregularSeal1
    /// </summary>
    IrregularSeal1,

    /// <summary>
    ///     IrregularSeal2
    /// </summary>
    IrregularSeal2,

    /// <summary>
    ///     FoldedCorner
    /// </summary>
    FoldedCorner,

    /// <summary>
    ///     Bevel
    /// </summary>
    Bevel,

    /// <summary>
    ///     Frame
    /// </summary>
    Frame,

    /// <summary>
    ///     HalfFrame
    /// </summary>
    HalfFrame,

    /// <summary>
    ///     Corner
    /// </summary>
    Corner,

    /// <summary>
    ///     DiagonalStripe
    /// </summary>
    DiagonalStripe,

    /// <summary>
    ///     Chord
    /// </summary>
    Chord,

    /// <summary>
    ///     Arc
    /// </summary>
    Arc,

    /// <summary>
    ///     LeftBracket
    /// </summary>
    LeftBracket,

    /// <summary>
    ///     RightBracket
    /// </summary>
    RightBracket,

    /// <summary>
    ///     LeftBrace
    /// </summary>
    LeftBrace,

    /// <summary>
    ///     RightBrace
    /// </summary>
    RightBrace,

    /// <summary>
    ///     BracketPair
    /// </summary>
    BracketPair,

    /// <summary>
    ///     BracePair
    /// </summary>
    BracePair,

    /// <summary>
    ///     StraightConnector1
    /// </summary>
    StraightConnector1,

    /// <summary>
    ///     BentConnector2
    /// </summary>
    BentConnector2,

    /// <summary>
    ///     BentConnector3
    /// </summary>
    BentConnector3,

    /// <summary>
    ///     BentConnector4
    /// </summary>
    BentConnector4,

    /// <summary>
    ///     BentConnector5
    /// </summary>
    BentConnector5,

    /// <summary>
    ///     CurvedConnector2
    /// </summary>
    CurvedConnector2,

    /// <summary>
    ///     CurvedConnector3
    /// </summary>
    CurvedConnector3,

    /// <summary>
    ///     CurvedConnector4
    /// </summary>
    CurvedConnector4,

    /// <summary>
    ///     CurvedConnector5
    /// </summary>
    CurvedConnector5,

    /// <summary>
    ///     Callout1
    /// </summary>
    Callout1,

    /// <summary>
    ///     Callout2
    /// </summary>
    Callout2,

    /// <summary>
    ///     Callout3
    /// </summary>
    Callout3,

    /// <summary>
    ///     AccentCallout1
    /// </summary>
    AccentCallout1,

    /// <summary>
    ///     AccentCallout2
    /// </summary>
    AccentCallout2,

    /// <summary>
    ///     AccentCallout3
    /// </summary>
    AccentCallout3,

    /// <summary>
    ///     BorderCallout1
    /// </summary>
    BorderCallout1,

    /// <summary>
    ///     BorderCallout2
    /// </summary>
    BorderCallout2,

    /// <summary>
    ///     BorderCallout3
    /// </summary>
    BorderCallout3,

    /// <summary>
    ///     AccentBorderCallout1
    /// </summary>
    AccentBorderCallout1,

    /// <summary>
    ///     AccentBorderCallout2
    /// </summary>
    AccentBorderCallout2,

    /// <summary>
    ///     AccentBorderCallout3
    /// </summary>
    AccentBorderCallout3,

    /// <summary>
    ///     WedgeRectangleCallout
    /// </summary>
    WedgeRectangleCallout,

    /// <summary>
    ///     WedgeRoundRectangleCallout
    /// </summary>
    WedgeRoundRectangleCallout,

    /// <summary>
    ///     WedgeEllipseCallout
    /// </summary>
    WedgeEllipseCallout,

    /// <summary>
    ///     CloudCallout
    /// </summary>
    CloudCallout,

    /// <summary>
    ///     Cloud
    /// </summary>
    Cloud,

    /// <summary>
    ///     Ribbon
    /// </summary>
    Ribbon,

    /// <summary>
    ///     Ribbon2
    /// </summary>
    Ribbon2,

    /// <summary>
    ///     EllipseRibbon
    /// </summary>
    EllipseRibbon,

    /// <summary>
    ///     EllipseRibbon2
    /// </summary>
    EllipseRibbon2,

    /// <summary>
    ///     LeftRightRibbon
    /// </summary>
    LeftRightRibbon,

    /// <summary>
    ///     VerticalScroll
    /// </summary>
    VerticalScroll,

    /// <summary>
    ///     HorizontalScroll
    /// </summary>
    HorizontalScroll,

    /// <summary>
    ///     Wave
    /// </summary>
    Wave,

    /// <summary>
    ///     DoubleWave
    /// </summary>
    DoubleWave,

    /// <summary>
    ///     Plus
    /// </summary>
    Plus,

    /// <summary>
    ///     FlowChartProcess
    /// </summary>
    FlowChartProcess,

    /// <summary>
    ///     FlowChartDecision
    /// </summary>
    FlowChartDecision,

    /// <summary>
    ///     FlowChartInputOutput
    /// </summary>
    FlowChartInputOutput,

    /// <summary>
    ///     FlowChartPredefinedProcess
    /// </summary>
    FlowChartPredefinedProcess,

    /// <summary>
    ///     FlowChartInternalStorage
    /// </summary>
    FlowChartInternalStorage,

    /// <summary>
    ///     FlowChartDocument
    /// </summary>
    FlowChartDocument,

    /// <summary>
    ///     FlowChartMultidocument
    /// </summary>
    FlowChartMultidocument,

    /// <summary>
    ///     FlowChartTerminator
    /// </summary>
    FlowChartTerminator,

    /// <summary>
    ///     FlowChartPreparation
    /// </summary>
    FlowChartPreparation,

    /// <summary>
    ///     FlowChartManualInput
    /// </summary>
    FlowChartManualInput,

    /// <summary>
    ///     FlowChartManualOperation
    /// </summary>
    FlowChartManualOperation,

    /// <summary>
    ///     FlowChartConnector
    /// </summary>
    FlowChartConnector,

    /// <summary>
    ///     FlowChartPunchedCard
    /// </summary>
    FlowChartPunchedCard,

    /// <summary>
    ///     FlowChartPunchedTape
    /// </summary>
    FlowChartPunchedTape,

    /// <summary>
    ///     FlowChartSummingJunction
    /// </summary>
    FlowChartSummingJunction,

    /// <summary>
    ///     FlowChartOr
    /// </summary>
    FlowChartOr,

    /// <summary>
    ///     FlowChartCollate
    /// </summary>
    FlowChartCollate,

    /// <summary>
    ///     FlowChartSort
    /// </summary>
    FlowChartSort,

    /// <summary>
    ///     FlowChartExtract
    /// </summary>
    FlowChartExtract,

    /// <summary>
    ///     FlowChartMerge
    /// </summary>
    FlowChartMerge,

    /// <summary>
    ///     FlowChartOfflineStorage
    /// </summary>
    FlowChartOfflineStorage,

    /// <summary>
    ///     FlowChartOnlineStorage
    /// </summary>
    FlowChartOnlineStorage,

    /// <summary>
    ///     FlowChartMagneticTape
    /// </summary>
    FlowChartMagneticTape,

    /// <summary>
    ///     FlowChartMagneticDisk
    /// </summary>
    FlowChartMagneticDisk,

    /// <summary>
    ///     FlowChartMagneticDrum
    /// </summary>
    FlowChartMagneticDrum,

    /// <summary>
    ///     FlowChartDisplay
    /// </summary>
    FlowChartDisplay,

    /// <summary>
    ///     FlowChartDelay
    /// </summary>
    FlowChartDelay,

    /// <summary>
    ///     FlowChartAlternateProcess
    /// </summary>
    FlowChartAlternateProcess,

    /// <summary>
    ///     FlowChartOffpageConnector
    /// </summary>
    FlowChartOffpageConnector,

    /// <summary>
    ///     ActionButtonBlank
    /// </summary>
    ActionButtonBlank,

    /// <summary>
    ///     ActionButtonHome
    /// </summary>
    ActionButtonHome,

    /// <summary>
    ///     ActionButtonHelp
    /// </summary>
    ActionButtonHelp,

    /// <summary>
    ///     ActionButtonInformation
    /// </summary>
    ActionButtonInformation,

    /// <summary>
    ///     ActionButtonForwardNext
    /// </summary>
    ActionButtonForwardNext,

    /// <summary>
    ///     ActionButtonBackPrevious
    /// </summary>
    ActionButtonBackPrevious,

    /// <summary>
    ///     ActionButtonEnd
    /// </summary>
    ActionButtonEnd,

    /// <summary>
    ///     ActionButtonBeginning
    /// </summary>
    ActionButtonBeginning,

    /// <summary>
    ///     ActionButtonReturn
    /// </summary>
    ActionButtonReturn,

    /// <summary>
    ///     ActionButtonDocument
    /// </summary>
    ActionButtonDocument,

    /// <summary>
    ///     ActionButtonSound
    /// </summary>
    ActionButtonSound,

    /// <summary>
    ///     ActionButtonMovie
    /// </summary>
    ActionButtonMovie,

    /// <summary>
    ///     Gear6
    /// </summary>
    Gear6,

    /// <summary>
    ///     Gear9
    /// </summary>
    Gear9,

    /// <summary>
    ///     Funnel
    /// </summary>
    Funnel,

    /// <summary>
    ///     MathPlus
    /// </summary>
    MathPlus,

    /// <summary>
    ///     MathMinus
    /// </summary>
    MathMinus,

    /// <summary>
    ///     MathMultiply
    /// </summary>
    MathMultiply,

    /// <summary>
    ///     MathDivide
    /// </summary>
    MathDivide,

    /// <summary>
    ///     MathEqual
    /// </summary>
    MathEqual,

    /// <summary>
    ///     MathNotEqual
    /// </summary>
    MathNotEqual,

    /// <summary>
    ///     CornerTabs
    /// </summary>
    CornerTabs,

    /// <summary>
    ///     SquareTabs
    /// </summary>
    SquareTabs,

    /// <summary>
    ///     PlaqueTabs
    /// </summary>
    PlaqueTabs,

    /// <summary>
    ///     ChartX
    /// </summary>
    ChartX,

    /// <summary>
    ///     ChartStar
    /// </summary>
    ChartStar,

    /// <summary>
    ///     ChartPlus
    /// </summary>
    ChartPlus,

    /// <summary>
    ///     Custom
    /// </summary>
    Custom
}