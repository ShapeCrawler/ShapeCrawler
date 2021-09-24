using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Statics;
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

        public int X => PixelConverter.HorizontalEmuToPixel(aOffset.X);

        public int Y => PixelConverter.VerticalEmuToPixel(aOffset.Y);

        public int Width => PixelConverter.HorizontalEmuToPixel(aExtents.Cx);

        public int Height => PixelConverter.HorizontalEmuToPixel(aExtents.Cy);

        public void SetX(int x)
        {
            aOffset.X.Value = x;
        }

        public void SetY(int y)
        {
            aOffset.Y.Value = y;
        }

        public void SetWidth(int w)
        {
            aExtents.Cx.Value = w;
        }

        public void SetHeight(int y)
        {
            aExtents.Cy.Value = y;
        }
    }
}