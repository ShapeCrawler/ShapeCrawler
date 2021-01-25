using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    public class AutoShape : IShape
    {
        private readonly P.Shape _pShape;

        public uint Id => _pShape.NonVisualShapeProperties.NonVisualDrawingProperties.Id;
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; }
        public GeometryType GeometryType { get; }
        public Placeholder Placeholder { get; }
    }

    /// <summary>
    /// Represents an auto shape on a Slide Master.
    /// </summary>
    public class MasterAutoShape : MasterShape, IAutoShape
    {
        internal ISlide Slide { get; }

        public MasterAutoShape(SlideMasterSc slideMaster, P.Shape pShape) : base(pShape)
        {
            Slide = slideMaster;
        }

        public TextBoxSc TextBox => GetTextBox();

        private TextBoxSc GetTextBox()
        {
            P.TextBody pTextBody = CompositeElement.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return new TextBoxSc(this);
            }

            return new TextBoxSc(this, pTextBody);
        }
    }
}