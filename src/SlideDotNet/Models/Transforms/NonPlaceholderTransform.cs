using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.Transforms
{
    public class NonPlaceholderTransform : ILocation
    {
        private readonly A.Offset _offset;

        private readonly A.Extents _extents;

        public long X => _offset.X.Value;

        public long Y => _offset.Y.Value;

        public long Width => _extents.Cx.Value;

        public long Height => _extents.Cy.Value;

        public NonPlaceholderTransform(OpenXmlCompositeElement sdkCompositeElement)
        {
            _offset = sdkCompositeElement.Descendants<A.Offset>().First(); //TODO: make lazy
            _extents = sdkCompositeElement.Descendants<A.Extents>().First();
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