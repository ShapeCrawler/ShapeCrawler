using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;

namespace SlideDotNet.Models.Transforms
{
    public class NonPlaceholderTransform : IInnerTransform
    {
        private readonly DocumentFormat.OpenXml.Drawing.Offset _offset;

        private readonly DocumentFormat.OpenXml.Drawing.Extents _extents;

        public long X => _offset.X.Value;

        public long Y => _offset.Y.Value;

        public long Width => _extents.Cx.Value;

        public long Height => _extents.Cy.Value;

        public NonPlaceholderTransform(OpenXmlCompositeElement xmlElement)
        {
            _offset = xmlElement.Descendants<DocumentFormat.OpenXml.Drawing.Offset>().First(); //TODO: make lazy
            _extents = xmlElement.Descendants<DocumentFormat.OpenXml.Drawing.Extents>().First();
        }

        public void SetX(long x) // TODO: validate
        {
            _offset.X.Value = x; 
        }

        public void SetY(long y)
        {
            _offset.Y.Value = y;
        }

        public void SetWidth(long w)
        {
            _extents.Cx.Value = w;
        }

        public void SetHeight(long y)
        {
            _extents.Cy.Value = y;
        }
    }
}