using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Exceptions;
using SlideDotNet.Models.SlideComponents;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.Transforms
{
    public class NonPlaceholderGroupedTransform : IInnerTransform
    {
        public long X { get; }

        public long Y { get; }

        public long Width { get; }

        public long Height { get; }

        public NonPlaceholderGroupedTransform(OpenXmlCompositeElement xmlElement, P.GroupShape groupShape)
        {
            var offset = xmlElement.Descendants<A.Offset>().First(); //TODO: make lazy
            var transformGroup = groupShape.GroupShapeProperties.TransformGroup;
            X = offset.X - transformGroup.ChildOffset.X + transformGroup.Offset.X;
            Y = offset.Y - transformGroup.ChildOffset.Y + transformGroup.Offset.Y;

            var extents = xmlElement.Descendants<A.Extents>().First();
            Width = extents.Cx.Value;
            Height = extents.Cy.Value;
        }

        public void SetX(long x)
        {
            throw new SlideDotNetException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetY(long y)
        {
            throw new SlideDotNetException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetWidth(long w)
        {
            throw new SlideDotNetException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetHeight(long h)
        {
            throw new SlideDotNetException(ExceptionMessages.ForGroupedCanNotChanged);
        }
    }
}