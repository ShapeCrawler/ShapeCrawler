using System.Collections.Generic;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an auto shape on a Slide Master.
    /// </summary>
    public class MasterAutoShape : MasterShape, IAutoShape
    {
        public MasterAutoShape(SlideMasterSc slideMaster, P.Shape pShape) : base(slideMaster, pShape)
        {
        }

        public ShapeContentType ContentType { get; }
        public bool HasTextBox { get; }
        public bool HasChart { get; }
        public bool HasPicture { get; }
        public ChartSc Chart { get; }
        public PictureSc Picture { get; }
        public IList<GroupShapeSc> GroupedShapes { get; }
        public OLEObject OleObject { get; }
        public bool IsPlaceholder { get; }
        public bool IsGrouped { get; }
        public SlideSc Slide { get; }

        public TextBoxSc TextBox => GetTextBox();

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public Placeholder Placeholder { get; }
        public ShapeFill Fill { get; }
        public string CustomData { get; set; }

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