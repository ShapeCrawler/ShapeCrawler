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
            var pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return new SCTextBox(this);
            }

            var aTexts = pTextBody.Descendants<A.Text>();
            if (aTexts.Sum(t => t.Text.Length) > 0)
            {
                return new SCTextBox( this, pTextBody);
            }

            return null;
        }

        private ShapeFill GetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
        {
            return new ShapeFill(this);
        }
    }
}