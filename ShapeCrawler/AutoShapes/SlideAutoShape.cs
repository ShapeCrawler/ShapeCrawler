using System;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents AutoShape located on Slide.
    /// </summary>
    internal class SlideAutoShape : SlideShape, IAutoShape, ITextFrameContainer
    {
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<TextFrame?> textFrame;

        internal SlideAutoShape(P.Shape pShape, SCSlide slideInternal, SlideGroupShape groupShape)
            : base(pShape, slideInternal, groupShape)
        {
            this.textFrame = new Lazy<TextFrame?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.GetFill);
        }

        #region Public Properties

        public IShapeFill Fill => this.shapeFill.Value;

        public Shape Shape => this; // TODO: should be internal?

        public SCShapeType ShapeType => SCShapeType.AutoShape;
        
        public ITextFrame? TextFrame => this.textFrame.Value;

        #endregion Public Properties

        private TextFrame? GetTextBox()
        {
            var pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            var canChangeTextFrame = this.Placeholder is { Type: PlaceholderType.Title } or null;
            return pTextBody == null ? null : new TextFrame(this, pTextBody, canChangeTextFrame);
        }

        private ShapeFill GetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            return new ShapeFill(this);
        }
    }
}