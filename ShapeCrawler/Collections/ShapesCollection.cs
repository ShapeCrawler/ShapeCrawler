using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Charts;
using ShapeCrawler.Factories;
using ShapeCrawler.Factories.Placeholders;
using ShapeCrawler.Factories.ShapeCreators;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.Tables;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a collection of the slide shapes.
    /// </summary>
    public class ShapesCollection : LibraryCollection<ShapeSc>
    {
        #region Constructors

        internal ShapesCollection(List<ShapeSc> shapes)
        {
            CollectionItems = shapes;
        }

        #endregion Constructors

        internal static ShapesCollection CreateForUserSlide(SlidePart slidePart, PresentationData preData, SlideSc slide)
        {
            var sldPhFontService = new PlaceholderFontService(slidePart);
            var phService = new PlaceholderService(slidePart.SlideLayoutPart);
            var transformFactory = new LocationParser(phService);
            var geometryFactory = new GeometryFactory(phService);
            var shapeContextBuilder = new ShapeContext.Builder(preData, sldPhFontService, slidePart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var pShapeHandler = new PShapeHandler(shapeContextBuilder, transformFactory, geometryFactory);
            var pictureHandler = new PictureHandler(shapeContextBuilder, transformFactory, geometryFactory, slidePart);
            var sdkGroupShapeHandler = new SdkGroupShapeHandler(shapeContextBuilder, transformFactory, geometryFactory, slidePart);

            pShapeHandler.Successor = sdkGroupShapeHandler;
            sdkGroupShapeHandler.Successor = oleGrFrameHandler;
            // OLE objects handler must be before pictures handler, cause OLE container can contain p:pic elements,
            // thereby defining OLE as a picture
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var shapes = new List<ShapeSc>(shapeTree.Count());
            foreach (OpenXmlElement openXmlElement in shapeTree)
            {
                ShapeSc shape = pShapeHandler.Create(openXmlElement);
                if (shape != null)
                {
                    shape.Slide = slide;
                    shapes.Add(shape);
                }
            }

            return new ShapesCollection(shapes);
        }

        public static MasterShapesCollection CreateForMasterSlide(ShapeTree shapeTree)
        {
            var slideMasterShapes = new List<BaseShape>();
            foreach (OpenXmlElement openXmlElement in shapeTree)
            {
                if (openXmlElement is P.Shape pShape)
                {
                    slideMasterShapes.Add(new MasterAutoShape(pShape));
                    continue;
                }
                
                if (openXmlElement is P.GraphicFrame pGraphicFrame)
                {
                    A.GraphicData aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                    if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole", StringComparison.Ordinal))
                    {
                        slideMasterShapes.Add(new OleObjectSc(pGraphicFrame));
                    }
                    else if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart", StringComparison.Ordinal))
                    {
                        slideMasterShapes.Add(new ChartSc(pGraphicFrame));
                    }
                    else if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table", StringComparison.Ordinal))
                    {
                        slideMasterShapes.Add(new TableSc(pGraphicFrame));
                        continue;
                    }
                }

                P.Picture pPicture;
                if (openXmlElement is P.Picture treePicture)
                {
                    pPicture = treePicture;
                }
                else
                {
                    P.Picture framePicture = openXmlElement.Descendants<P.Picture>().FirstOrDefault();
                    pPicture = framePicture;
                }
                if (pPicture != null)
                {
                    slideMasterShapes.Add(new PictureSc(pPicture));
                }
            }

            return new MasterShapesCollection(slideMasterShapes);
        }
    }
}