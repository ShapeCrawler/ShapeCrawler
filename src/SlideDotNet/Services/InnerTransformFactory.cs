using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.Transforms;
using SlideDotNet.Services.Placeholders;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Services
{
    public class InnerTransformFactory
    {
        private readonly IPlaceholderService _phService;

        public InnerTransformFactory(IPlaceholderService phService)
        {
            _phService = phService;
        }

        public IInnerTransform FromComposite(OpenXmlCompositeElement sdkCompositeElement)
        {
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
                var placeholderLocationData = _phService.TryGet(sdkCompositeElement);
                innerTransform = new PlaceholderTransform(placeholderLocationData);
            }

            return innerTransform;
        }
    }
}