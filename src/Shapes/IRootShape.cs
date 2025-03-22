namespace ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

/// <summary>
///     Root (non-grouped) Auto Shape.
/// </summary>
internal interface IRootShape : IShape
{
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    void Duplicate();
}

internal sealed class RootShape : CopyableShape, IRootShape
{
    private readonly IShape decoratedShape;
    private readonly P.Shape pShape;

    internal RootShape(
        P.Shape pShape,
        IShape decoratedShape)
        : base(pShape)
    {
        this.decoratedShape = decoratedShape;
        this.pShape = pShape;
    }

    #region Decorated Shape

    public override ShapeType ShapeType => this.decoratedShape.ShapeType;

    public override bool HasOutline => this.decoratedShape.HasOutline;

    public override IShapeOutline Outline => this.decoratedShape.Outline;

    public override bool HasFill => this.decoratedShape.HasFill;

    public override IShapeFill Fill => this.decoratedShape.Fill;

    public override bool IsTextHolder => this.decoratedShape.IsTextHolder;

    public override ITextBox TextBox => this.decoratedShape.TextBox;

    public override Geometry GeometryType
    {
        get => this.decoratedShape.GeometryType;
        set => this.decoratedShape.GeometryType = value;
    }

    public override decimal CornerSize
    {
        get => this.decoratedShape.CornerSize;
        set => this.decoratedShape.CornerSize = value;
    }

    public override decimal X
    {
        get => this.decoratedShape.X;
        set => this.decoratedShape.X = value;
    }

    public override decimal[] Adjustments
    {
        get => this.decoratedShape.Adjustments;
        set => this.decoratedShape.Adjustments = value;
    }

    #endregion Decorated Shape

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShape.Parent!;
        var autoShapes = new SCPShapeTree(pShapeTree);
        autoShapes.Add(this.pShape);
    }
}