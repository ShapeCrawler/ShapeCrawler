using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories
{
    internal class NonPlaceholderTransform : ILocation
    {
        private readonly A.Extents aExtents;
        private readonly A.Offset aOffset;

        public NonPlaceholderTransform(OpenXmlCompositeElement sdkCompositeElement)
        {
            this.aOffset = sdkCompositeElement.Descendants<A.Offset>().First();
            this.aExtents = sdkCompositeElement.Descendants<A.Extents>().First();
        }

        public long X => aOffset.X.Value;

        public long Y => aOffset.Y.Value;

        public long Width => aExtents.Cx.Value;

        public long Height => aExtents.Cy.Value;

        public void SetX(long x)
        {
            aOffset.X.Value = x;
        }

        public void SetY(long y)
        {
            aOffset.Y.Value = y;
        }

        public void SetWidth(long w)
        {
            aExtents.Cx.Value = w;
        }

        public void SetHeight(long y)
        {
            aExtents.Cy.Value = y;
        }
    }
}