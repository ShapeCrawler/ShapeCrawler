using System;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing
{
    internal class LayoutPicture : LayoutShape, IShape
    {
        public LayoutPicture(SlideLayoutSc slideLayout, P.Picture pPicture) : base(slideLayout, pPicture)
        {
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

        public string Name => throw new NotImplementedException();

        public bool Hidden => throw new NotImplementedException();

        public IPlaceholder Placeholder => throw new NotImplementedException();

        public string CustomData
        {
            get => throw new NotImplementedException();
            set => throw new NotImplementedException();
        }
    }
}