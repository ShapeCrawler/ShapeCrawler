using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Audio;
using ShapeCrawler.Drawing;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            P.Picture pPicture;
            if (pShapeTreeChild is P.Picture treePic)
            {
                A.AudioFromFile aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                    .GetFirstChild<A.AudioFromFile>();
                if (aAudioFile != null)
                {
                    return new AudioShape(slide, pShapeTreeChild);
                }

                pPicture = treePic;
            }
            else
            {
                pPicture = pShapeTreeChild.Descendants<P.Picture>().FirstOrDefault();
            }

            if (pPicture == null)
            {
                return this.Successor?.Create(pShapeTreeChild, slide);
            }

            StringValue picReference = pPicture.GetFirstChild<P.BlipFill>()?.Blip?.Embed;
            if (picReference == null)
            {
                return null;
            }

            SlidePicture picture = new (slide, pPicture, picReference);

            return picture;
        }

        public override IShape CreateGroupedShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            throw new NotImplementedException();
        }
    }
}