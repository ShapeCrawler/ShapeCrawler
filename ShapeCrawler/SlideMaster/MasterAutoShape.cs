using ShapeCrawler.Models;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
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
            P.TextBody pTextBody = _compositeElement.GetFirstChild<P.TextBody>();
            if (pTextBody == null)
            {
                return new TextBoxSc(this);
            }

            return new TextBoxSc(this, pTextBody);
        }
    }
}