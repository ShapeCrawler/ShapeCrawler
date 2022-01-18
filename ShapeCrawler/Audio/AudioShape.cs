using DocumentFormat.OpenXml;

namespace ShapeCrawler.Audio
{
    internal class AudioShape : SlideShape, IAudioShape
    {
        public AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlideInternal)
            : base(pShapeTreesChild, parentSlideInternal, null)
        {
        }

        public byte[] BinaryData { get; }
    }
}