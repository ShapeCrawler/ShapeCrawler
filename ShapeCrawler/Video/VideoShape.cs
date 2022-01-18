using DocumentFormat.OpenXml;

namespace ShapeCrawler.Video
{
    internal class VideoShape : SlideShape, IVideoShape
    {
        public VideoShape(SCSlide parentSlideInternal, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(sdkPShapeTreeChild, parentSlideInternal, null)
        {
        }

        public byte[] BinaryData { get; }
    }
}
