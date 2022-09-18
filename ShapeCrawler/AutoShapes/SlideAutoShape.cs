using System;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
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
        private readonly Lazy<TextFrame?> textBox;

        internal SlideAutoShape(P.Shape pShape, SCSlide slideInternal, SlideGroupShape groupShape)
            : base(pShape, slideInternal, groupShape)
        {
            this.textBox = new Lazy<TextFrame?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.GetFill);
        }

        #region Public Properties

        public ITextFrame TextFrame => this.textBox.Value;

        public IShapeFill Fill => this.shapeFill.Value;

        public IShape Shape => this; // TODO: should be internal?

        public ShapeType ShapeType => ShapeType.AutoShape;

        #endregion Public Properties

        private TextFrame? GetTextBox()
        {
            var pTxBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            return pTxBody == null ? null : new TextFrame(this, pTxBody);
        }

        private ShapeFill GetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            return new ShapeFill(this);
        }
    }
}