using DocumentFormat.OpenXml;

namespace ShapeCrawler.Audio
{
    internal class AudioShape : SlideShape, IAudioShape
    {
        public AudioShape(SCSlide parentSlide, OpenXmlCompositeElement sdkPShapeTreeChild)
            : base(parentSlide, sdkPShapeTreeChild)
        {
        }

        public byte[] BinaryData { get; }
    }
}