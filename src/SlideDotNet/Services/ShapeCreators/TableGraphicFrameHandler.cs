using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Models.TableComponents;
using SlideDotNet.Services.Builders;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Services.ShapeCreators
{
    public class TableGraphicFrameHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder _shapeContextBuilder;
        private readonly InnerTransformFactory _transformFactory;
        private readonly IShapeBuilder _shapeBuilder;
        private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/table";

        #region Constructors

        public TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder, InnerTransformFactory transformFactory) :
            this(shapeContextBuilder, transformFactory, new ShapeEx.Builder())
        {
            
        }

        public TableGraphicFrameHandler(ShapeContext.Builder shapeContextBuilder,
                                        InnerTransformFactory transformFactory,
                                        IShapeBuilder shapeBuilder)
        {
            _shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
            _transformFactory = transformFactory ?? throw new ArgumentNullException(nameof(transformFactory));
            _shapeBuilder = shapeBuilder ?? throw new ArgumentNullException(nameof(shapeBuilder));
        }

        #endregion Constructors

        public override ShapeEx Create(OpenXmlElement sdkElement)
        {
            Check.NotNull(sdkElement, nameof(sdkElement));

            if (sdkElement is P.GraphicFrame sdkGraphicFrame)
            {
                var grData = sdkElement.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                if (grData.Uri.Value.Equals(Uri, StringComparison.Ordinal))
                {
                    var spContext = _shapeContextBuilder.Build(sdkElement);
                    var innerTransform = _transformFactory.FromComposite(sdkGraphicFrame);
                    var table = new TableEx(sdkGraphicFrame, spContext);
                    var shape = _shapeBuilder.WithTable(innerTransform, spContext, table);

                    return shape;
                }
            }

            if (Successor != null)
            {
                return Successor.Create(sdkElement);
            }

            return null;
        }
    }
}