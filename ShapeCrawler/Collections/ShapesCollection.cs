using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Factories;
using ShapeCrawler.Factories.Placeholders;
using ShapeCrawler.Factories.ShapeCreators;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;
using ShapeCrawler.SlideMaster;
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

        internal static ShapesCollection CreateForUserSlide(SlidePart slidePart, SlideSc slide)
        {
            var sldPhFontService = new PlaceholderFontService(slidePart);
            var phService = new PlaceholderService(slidePart.SlideLayoutPart);
            var transformFactory = new LocationParser(phService);
            var geometryFactory = new GeometryFactory(phService);
            var shapeContextBuilder = new ShapeContext.Builder(sldPhFontService, slidePart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var pShapeHandler = new PShapeHandler(shapeContextBuilder, transformFactory, geometryFactory);
            var pictureHandler = new PictureHandler(shapeContextBuilder, transformFactory, geometryFactory, slidePart);
            var sdkGroupShapeHandler = new PGroupShapeHandler(shapeContextBuilder, transformFactory, geometryFactory, slidePart);

            pShapeHandler.Successor = sdkGroupShapeHandler;
            sdkGroupShapeHandler.Successor = oleGrFrameHandler;
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var shapes = new List<ShapeSc>(shapeTree.Count());
            foreach (OpenXmlCompositeElement compositeElement in shapeTree.OfType<OpenXmlCompositeElement>())
            {
                ShapeSc shape = pShapeHandler.Create(compositeElement, slide);
                if (shape != null)
                {
                    shape.Slide = slide;
                    shapes.Add(shape);
                }
            }

            return new ShapesCollection(shapes);
        }

        public static MasterShapesCollection CreateForMasterSlide(SlideMasterSc slideMaster, ShapeTree shapeTree)
        {
            var slideMasterShapes = new List<BaseShape>();
            foreach (OpenXmlCompositeElement compositeElement in shapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        slideMasterShapes.Add(new MasterAutoShape(slideMaster, pShape));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData = pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole", StringComparison.Ordinal))
                        {
                            slideMasterShapes.Add(new MasterOLEObject(slideMaster, pGraphicFrame));
                            continue;
                        }
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart", StringComparison.Ordinal))
                        {
                            slideMasterShapes.Add(new ChartScNew(slideMaster, pGraphicFrame));
                            continue;
                        }
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table", StringComparison.Ordinal))
                        {
                            slideMasterShapes.Add(new TableNew(slideMaster, pGraphicFrame));
                            continue;
                        }

                        break;
                    }
                }

                // OLE Objects should be parsed before pictures, since OLE containers can contain p:pic elements,
                // thus OLE objects can be parsed as a picture by mistake.
                P.Picture pPicture;
                if (compositeElement is P.Picture treePicture)
                {
                    pPicture = treePicture;
                }
                else
                {
                    P.Picture framePicture = compositeElement.Descendants<P.Picture>().FirstOrDefault();
                    pPicture = framePicture;
                }
                if (pPicture != null)
                {
                    slideMasterShapes.Add(new PictureScNew(slideMaster, pPicture));
                }
            }

            return new MasterShapesCollection(slideMasterShapes);
        }
    }
}