using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;

namespace ShapeCrawler.Media
{
    internal abstract class MediaShape : SlideShape
    {
        protected MediaShape(OpenXmlCompositeElement childOfPShapeTree, SCSlide slide, Shape groupShape) 
            : base(childOfPShapeTree, slide, groupShape)
        {
        }

        public byte[] BinaryData => this.GetBinaryData();

        private byte[] GetBinaryData()
        {
            var pPic = (DocumentFormat.OpenXml.Presentation.Picture)this.PShapeTreesChild;
            var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship = this.Slide.SDKSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media!.Embed!.Value);
            var stream = relationship.DataPart.GetStream();
            var bytes = stream.ToArray();
            stream.Close();

            return bytes;
        }
    }
}