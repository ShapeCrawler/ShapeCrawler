using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OleObject = SlideDotNet.Models.SlideComponents.OleObject;

namespace SlideDotNet.Services.ShapeCreators
{
    public class OleGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly IPreSettings _preSettings;
        private readonly SlidePlaceholderFontService _sldFontService;
        private readonly SlidePart _sdkSldPart;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        private const string Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

        public OleGraphicFrameHandler(IPreSettings preSettings,
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

            if (openXmlElement is P.GraphicFrame sdkGraphicFrame)
            {
                var grData = openXmlElement.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri))
                {
                    var spContext = new ShapeContext(_preSettings, _sldFontService, openXmlElement, _sdkSldPart);
                    var innerTransform = _transformFactory.FromComposite(sdkGraphicFrame);
                    var ole = new OleObject(sdkGraphicFrame);
                    var shape = _shapeBuilder.WithOle(innerTransform, spContext, ole);

                    return shape;
                }
            }

            if (Successor != null)
            {
                return Successor.Create(openXmlElement);
            }

            return null;
        }
    }
}