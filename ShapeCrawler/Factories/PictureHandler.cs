using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Video;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents handler for p:pic and p:graphicFrame elements.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        internal override IShape? Create(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            P.Picture? pPicture;
            if (pShapeTreesChild is P.Picture treePic)
            {
                OpenXmlElement element = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.ChildElements.FirstOrDefault();

                switch (element)
                {
                    case AudioFromFile:
                    {
                        A.AudioFromFile? aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                            .GetFirstChild<A.AudioFromFile>();
                        if (aAudioFile is not null)
                        {
                            return new AudioShape(pShapeTreesChild, slide);
                        }

                        break;
                    }
                    case VideoFromFile file:
                    {
                        A.VideoFromFile aVideoFile = file;

                        if (aVideoFile != null)
                        {
                            return new VideoShape(slide, pShapeTreesChild);
                        }

                        break;
                    }
                }

                pPicture = treePic;
            }
            else
            {
                pPicture = pShapeTreesChild.Descendants<P.Picture>().FirstOrDefault();
            }

            if (pPicture == null)
            {
                return this.Successor?.Create(pShapeTreesChild, slide, groupShape);
            }

            StringValue? picReference = pPicture.GetFirstChild<P.BlipFill>()?.Blip?.Embed;
            if (picReference is null)
            {
                return null;
            }

            SlidePicture picture = new (pPicture, slide, picReference);

            return picture;
        }
    }
}