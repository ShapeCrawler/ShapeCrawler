using System;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    internal class LayoutPicture : IShape
    {
        public LayoutPicture(SlideLayoutSc slideLayout, DocumentFormat.OpenXml.Presentation.Picture pPicture)
        {
            throw new NotImplementedException();
        }

        public long X
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        public long Y
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        public long Width
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        public long Height
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }

        public int Id => throw new NotImplementedException();

        public string Name => throw new NotImplementedException();

        public bool Hidden => throw new NotImplementedException();

        public IPlaceholder Placeholder => throw new NotImplementedException();

        public GeometryType GeometryType => throw new NotImplementedException();

        public string CustomData
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }
    }
}