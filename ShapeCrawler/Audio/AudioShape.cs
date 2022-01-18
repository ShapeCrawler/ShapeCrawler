using DocumentFormat.OpenXml;

namespace ShapeCrawler.Audio
{
    internal class AudioShape : SlideShape, IAudioShape
    {
        public AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlideLayoutInternal)
            : base(pShapeTreesChild, parentSlideLayoutInternal, null)
        {
        }

        public byte[] BinaryData { get; }
    }
}