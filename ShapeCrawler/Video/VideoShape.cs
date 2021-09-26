using DocumentFormat.OpenXml;

namespace ShapeCrawler.Video
{
    internal class VideoShape : SlideShape, IVideoShape
    {
        public VideoShape(SCSlide parentSlide, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(parentSlide, sdkPShapeTreeChild)
        {
        }

        public byte[] BinaryData { get; }
    }
}
