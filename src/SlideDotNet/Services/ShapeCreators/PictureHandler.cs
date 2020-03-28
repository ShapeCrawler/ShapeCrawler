using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.ShapeCreators
{
    /// <summary>
    /// Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    public class PictureHandler : OpenXmlElementHandler
    {
        private readonly IPreSettings _preSettings;
        private readonly SlidePlaceholderFontService _sldFontService;
        private readonly SlidePart _sdkSldPart;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        public PictureHandler(IPreSettings preSettings,
            SlidePlaceholderFontService sldFontService,
            SlidePart sdkSldPart,
            InnerTransformFactory transformFactory,
            IShapeBuilder shapeBuilder)
        {
            _preSettings = preSettings ?? throw new ArgumentNullException(nameof(preSettings));
            _sldFontService = sldFontService ?? throw new ArgumentNullException(nameof(sldFontService));
            _sdkSldPart = sdkSldPart ?? throw new ArgumentNullException(nameof(sdkSldPart));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder;
        }

        public override ShapeEx Create(OpenXmlElement openXmlElement)
        {
            Check.NotNull(openXmlElement, nameof(openXmlElement));

            P.Picture sdkPicture;
            if (openXmlElement is P.Picture treePic)
            {
                sdkPicture = treePic;
            }
            else
            {
                var framePic = openXmlElement.Descendants<P.Picture>().FirstOrDefault();
                sdkPicture = framePic;
            }
            if (sdkPicture != null)
            {
                var pBlipFill = sdkPicture.GetFirstChild<P.BlipFill>();
                var blipRelateId = pBlipFill?.Blip?.Embed?.Value;
                if (blipRelateId == null)
                {
                    return null;
                }
                var pictureEx = new PictureEx(_sdkSldPart, blipRelateId);
                var spContext = new ShapeContext(_preSettings, _sldFontService, openXmlElement, _sdkSldPart);
                var innerTransform = _transformFactory.FromComposite(sdkPicture);
                var shape = _shapeBuilder.WithPicture(innerTransform, spContext, pictureEx);

                return shape;

            }

            if (Successor != null)
            {
                return Successor.Create(openXmlElement);
            }

            return null;
        }
    }
}