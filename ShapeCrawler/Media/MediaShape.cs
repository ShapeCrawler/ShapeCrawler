using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Media
{
    internal abstract class MediaShape : SlideShape
    {
        protected MediaShape(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide, Shape? groupShape) 
            : base(pShapeTreeChild, slide, groupShape)
        {
        }

        public byte[] BinaryData => this.GetBinaryData();

        public string MIME => this.GetMime();

        private byte[] GetBinaryData()
        {
            var pPic = (P.Picture)this.PShapeTreesChild;
            var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship = this.Slide.SDKSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
            var stream = relationship.DataPart.GetStream();
            var bytes = stream.ToArray();
            stream.Close();

            return bytes;
        }
        
        private string GetMime()
        {
            var pPic = (P.Picture)this.PShapeTreesChild;
            var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship = this.Slide.SDKSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

            return relationship.DataPart.ContentType;
        }
    }
}