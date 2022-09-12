using System;
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
    internal class SlideAutoShape : SlideShape, IAutoShape, ITextBoxContainer
    {
        private readonly Lazy<ShapeFill> shapeFill;
        private readonly Lazy<SCTextBox?> textBox;

        internal SlideAutoShape(P.Shape pShape, SCSlide slideInternal, SlideGroupShape groupShape)
            : base(pShape, slideInternal, groupShape)
        {
            this.textBox = new Lazy<SCTextBox?>(this.GetTextBox);
            this.shapeFill = new Lazy<ShapeFill>(this.GetFill);
        }

        #region Public Properties

        public ITextBox TextBox => this.textBox.Value;

        public IShapeFill Fill => this.shapeFill.Value;

        public IShape Shape => this; // TODO: should be internal?

        public ShapeType ShapeType => ShapeType.AutoShape;

        #endregion Public Properties

        private SCTextBox GetTextBox()
        {
            var pTxBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            return pTxBody == null ? new SCTextBox(this) : new SCTextBox(this, pTxBody);
        }

        private ShapeFill GetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            return new ShapeFill(this);
        }
    }
}