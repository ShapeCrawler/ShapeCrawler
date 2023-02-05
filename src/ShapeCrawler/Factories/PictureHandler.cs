using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OneOf;
using ShapeCrawler.Media;
using ShapeCrawler.Pictures;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories;

internal sealed class PictureHandler : OpenXmlElementHandler
{
    internal override SCShape? Create(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection)
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
                        return new SCAudioShape(pShapeTreeChild, slideObject, shapeCollection);
                    }

                    break;
                }

                case VideoFromFile:
                {
                    return new SCVideoShape(pShapeTreeChild, slideObject, shapeCollection);
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
            return this.Successor?.Create(pShapeTreeChild, slideObject, shapeCollection);
        }

        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        var blipEmbed = aBlip?.Embed;
        if (blipEmbed is null)
        {
            return null;
        }

        var picture = new SCPicture(pPicture, slideObject, shapeCollection, aBlip!);

        return picture;
    }
}