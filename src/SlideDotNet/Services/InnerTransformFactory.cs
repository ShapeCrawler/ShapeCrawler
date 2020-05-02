using System;
using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.Transforms;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    public class InnerTransformFactory
    {
        #region Fields

        private readonly IPlaceholderService _phService;

        #endregion Fields

        #region Constructors

        public InnerTransformFactory(IPlaceholderService phService)
        {
            _phService = phService ?? throw new ArgumentNullException(nameof(phService));
        }

        #endregion Constructors

        #region Public Methods

        public IInnerTransform FromComposite(OpenXmlCompositeElement sdkCompositeElement)
        {
            Check.NotNull(sdkCompositeElement, nameof(sdkCompositeElement));

            IInnerTransform innerTransform;
            var t2d = sdkCompositeElement.Descendants<A.Transform2D>().FirstOrDefault();
            if (t2d != null)
            {
                // Group
                if (sdkCompositeElement.Parent is P.GroupShape groupShape)
                {
                    innerTransform = new NonPlaceholderGroupedTransform(sdkCompositeElement, groupShape);
                }
                // ShapeTree
                else
                {
                    innerTransform = new NonPlaceholderTransform(sdkCompositeElement);
                }
            }
            else
            {
                var placeholderLocationData = _phService.TryGetLocation(sdkCompositeElement);
                innerTransform = new PlaceholderTransform(placeholderLocationData);
            }

            return innerTransform;
        }

        #endregion Public Methods
    }
}