using DocumentFormat.OpenXml;

namespace ShapeCrawler.Video
{
    internal class VideoShape : SlideShape, IVideoShape
    {
        public VideoShape(SCSlide parentSlide, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlide, null)
        {
        }

        public byte[] BinaryData { get; }
    }
}
