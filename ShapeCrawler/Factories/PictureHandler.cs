using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Drawing;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using OneOf;

namespace ShapeCrawler.Factories;

internal class PictureHandler : OpenXmlElementHandler
{
    internal override Shape? Create(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject, SCGroupShape groupShape)
    {
        P.Picture? pPicture;
        if (pShapeTreeChild is P.Picture treePic)
        {
            var element = treePic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.ChildElements.FirstOrDefault();

            switch (element)
            {
                case AudioFromFile:
                {
                    var aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                        .GetFirstChild<A.AudioFromFile>();
                    if (aAudioFile is not null)
                    {
                        return new AudioShape(pShapeTreeChild, slideObject);
                    }

                    break;
                }

                case VideoFromFile file:
                {
                    A.VideoFromFile aVideoFile = file;

                    if (aVideoFile != null)
                    {
                        return new VideoShape(slideObject, pShapeTreeChild);
                    }

                    break;
                }
            }

            pPicture = treePic;
        }
        else
        {
            pPicture = pShapeTreeChild.Descendants<P.Picture>().FirstOrDefault();
        }

        if (pPicture == null)
        {
            return this.Successor?.Create(pShapeTreeChild, slideObject, groupShape);
        }

        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        var blipEmbed = aBlip?.Embed;
        if (blipEmbed is null)
        {
            return null;
        }

        var picture = new SlidePicture(pPicture, slideObject, aBlip!);

        return picture;
    }
}