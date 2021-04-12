using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories
{
    internal class NonPlaceholderGroupedTransform : ILocation
    {
        public NonPlaceholderGroupedTransform(OpenXmlCompositeElement xmlElement, P.GroupShape groupShape)
        {
            var offset = xmlElement.Descendants<A.Offset>().First();
            var transformGroup = groupShape.GroupShapeProperties.TransformGroup;
            X = offset.X - transformGroup.ChildOffset.X + transformGroup.Offset.X;
            Y = offset.Y - transformGroup.ChildOffset.Y + transformGroup.Offset.Y;

            var extents = xmlElement.Descendants<A.Extents>().First();
            Width = extents.Cx.Value;
            Height = extents.Cy.Value;
        }

        public long X { get; }

        public long Y { get; }

        public long Width { get; }

        public long Height { get; }

        public void SetX(long x)
        {
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged); // TODO: add implementation
        }

        public void SetY(long y)
        {// TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged); 
        }

        public void SetWidth(long w)
        {// TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetHeight(long h)
        {// TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged);
        }
    }
}