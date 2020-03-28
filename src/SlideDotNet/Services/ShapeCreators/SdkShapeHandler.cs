using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services.ShapeCreators
{
    public class SdkShapeHandler : OpenXmlElementHandler
    {
        private readonly IPreSettings _preSettings;
        private readonly SlidePlaceholderFontService _sldFontService;
        private readonly SlidePart _sdkSldPart;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;

        public SdkShapeHandler(IPreSettings preSettings, 
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

            if (openXmlElement is P.Shape sdkShape)
            {
                var spContext = new ShapeContext(_preSettings, _sldFontService, sdkShape, _sdkSldPart);
                var innerTransform = _transformFactory.FromComposite(sdkShape);
                var shape = _shapeBuilder.WithAutoShape(innerTransform, spContext);
                
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