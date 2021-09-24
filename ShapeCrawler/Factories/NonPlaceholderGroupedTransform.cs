using System;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal class NonPlaceholderGroupedTransform : ILocation
    {
        public NonPlaceholderGroupedTransform(OpenXmlCompositeElement xmlElement, P.GroupShape groupShape)
        {
            A.Offset offset = xmlElement.Descendants<A.Offset>().First();
            A.TransformGroup transformGroup = groupShape.GroupShapeProperties.TransformGroup;
            this.X = PixelConverter.HorizontalEmuToPixel(offset.X - transformGroup.ChildOffset.X + transformGroup.Offset.X);
            this.Y = PixelConverter.VerticalEmuToPixel(offset.Y - transformGroup.ChildOffset.Y + transformGroup.Offset.Y);

            A.Extents extents = xmlElement.Descendants<A.Extents>().First();
            this.Width = PixelConverter.HorizontalEmuToPixel(extents.Cx.Value);
            this.Height = PixelConverter.VerticalEmuToPixel(extents.Cy.Value);
        }

        public int X { get; }

        public int Y { get; }

        public int Width { get; }

        public int Height { get; }

        public void SetX(int x)
        {
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged); // TODO: add implementation
        }

        public void SetY(int y)
        {
            // TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetWidth(int w)
        {
            // TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged);
        }

        public void SetHeight(int h)
        {
            // TODO: add implementation
            throw new ShapeCrawlerException(ExceptionMessages.ForGroupedCanNotChanged);
        }
    }

    internal class PixelConverter
    {
        private static readonly Bitmap Bitmap = new Bitmap(100, 100);

        internal static int HorizontalEmuToPixel(long horizontalEmu)
        {
            return (int)(horizontalEmu * Bitmap.HorizontalResolution / 914400);
        }

        internal static int VerticalEmuToPixel(long verticalEmu)
        {
            return (int)(verticalEmu * Bitmap.VerticalResolution / 914400);
        }
    }
}