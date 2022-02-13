using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Video
{
    internal class VideoShape : SlideShape, IVideoShape
    {
        public VideoShape(SCSlide parentSlideLayoutInternal, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideLayoutInternal, null)
        {
        }

        public byte[] BinaryData { get; }

        public ShapeType ShapeType => ShapeType.VideoShape;
    }
}
