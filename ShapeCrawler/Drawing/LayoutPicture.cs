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

        public string Name => "throw new NotImplementedException()"; //TODO: Implement

        public bool Hidden => false;//TODO: throw new NotImplementedException()

        public IPlaceholder Placeholder => null; //TODO: throw new NotImplementedException()

        public string CustomData
        {
            get => null;//TODO: throw new NotImplementedException()
            set => throw new NotImplementedException();
        }
    }
}