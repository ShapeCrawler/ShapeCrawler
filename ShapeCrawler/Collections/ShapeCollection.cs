using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Experiment;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of the slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
    {
        #region Constructors

        internal ShapeCollection(List<IShape> shapes)
        {
            CollectionItems = shapes;
        }

        #endregion Constructors

        internal static ShapeCollection CreateForSlide(SlidePart slidePart, SlideSc slide)
        {
            var sldPhFontService = new PlaceholderFontService(slidePart);
            var phService = new PlaceholderService(slidePart.SlideLayoutPart);
            var transformFactory = new LocationParser(phService);
            var geometryFactory = new GeometryFactory(phService);
            var shapeContextBuilder = new ShapeContext.Builder(sldPhFontService, slidePart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var pShapeHandler = new AutoShapeCreator(shapeContextBuilder, transformFactory, geometryFactory);
            var pictureHandler = new PictureHandler(shapeContextBuilder, transformFactory, geometryFactory);
            var sdkGroupShapeHandler =
                new PGroupShapeHandler(shapeContextBuilder, transformFactory, geometryFactory, slidePart);

            pShapeHandler.Successor = sdkGroupShapeHandler;
            sdkGroupShapeHandler.Successor = oleGrFrameHandler;
            oleGrFrameHandler.Successor = pictureHandler;
            pictureHandler.Successor = chartGrFrameHandler;
            chartGrFrameHandler.Successor = tableGrFrameHandler;

            P.ShapeTree shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;
            var shapes = new List<IShape>(shapeTree.Count());
            foreach (OpenXmlCompositeElement compositeElement in shapeTree.OfType<OpenXmlCompositeElement>())
            {
                IShape shape = pShapeHandler.Create(compositeElement, slide);
                if (shape != null)
                {
                    shapes.Add(shape);
                }
            }

            return new ShapeCollection(shapes);
        }

        internal static ShapeCollection CreateForSlideLayout(P.ShapeTree pShapeTree, SlideLayoutSc slideLayout)
        {
            var shapeList = new List<IShape>();
            foreach (OpenXmlCompositeElement compositeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        shapeList.Add(new SlideAutoShape(slideLayout, pShape));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                        {
                            A.GraphicData aGraphicData =
                                pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                            if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                                StringComparison.Ordinal))
                            {
                                shapeList.Add(new SlideOLEObject(slideLayout, pGraphicFrame));
                                continue;
                            }

                            if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                                StringComparison.Ordinal))
                            {
                                shapeList.Add(new SlideChart(slideLayout, pGraphicFrame));
                                continue;
                            }

                            if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                                StringComparison.Ordinal))
                            {
                                shapeList.Add(new SlideTable(slideLayout, pGraphicFrame));
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
                    shapeList.Add(new SlidePicture(slideLayout, pPicture));
                }
            }

            return new ShapeCollection(shapeList);
        }

        internal static MasterShapeCollection CreateForSlideMaster(SlideMasterSc slideMaster, P.ShapeTree shapeTree)
        {
            var slideMasterShapes = new List<ChartScNew>();
            foreach (OpenXmlCompositeElement compositeElement in shapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        slideMasterShapes.Add(new ChartScNew(slideMaster, new P.GraphicFrame()));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            slideMasterShapes.Add(new ChartScNew(slideMaster, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            slideMasterShapes.Add(new ChartScNew(slideMaster, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
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
                    slideMasterShapes.Add(new ChartScNew(slideMaster, new P.GraphicFrame()));
                }
            }

            return new MasterShapeCollection(slideMasterShapes);
        }

        internal Shape GetShapeByPPlaceholderShape(P.PlaceholderShape pPlaceholderShapeParam)
        {
            Shape mappedShape = CollectionItems.OfType<Shape>().First(collectionShape =>
            {
                P.PlaceholderShape pPlaceholderShape = ((Placeholder)collectionShape.Placeholder).PPlaceholderShape;
                if (pPlaceholderShapeParam.Type != null && pPlaceholderShape.Type != null)
                {
                    return pPlaceholderShapeParam.Type.Equals(pPlaceholderShape.Type);
                }

                if (pPlaceholderShapeParam.Type == null && pPlaceholderShape.Type == null)
                {
                    return pPlaceholderShapeParam.Index == pPlaceholderShape.Index;
                }

                return false;
            });

            return mappedShape;
        }
    }

    internal class TableNew : ChartScNew
    {
        public TableNew(SlideMasterSc slideMasterSc, P.GraphicFrame graphicFrame) : base()
        {
            throw new NotImplementedException();
        }
    }

    internal class MasterOLEObject : Shape
    {
        private SlideMasterSc slideMaster;
        private P.GraphicFrame pGraphicFrame;

        public MasterOLEObject(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base()
        {
            this.slideMaster = slideMaster;
            this.pGraphicFrame = pGraphicFrame;
        }

        public override IPlaceholder Placeholder => throw new NotImplementedException();
    }

    internal class PictureScNew : Shape
    {
        private SlideMasterSc slideMaster;
        private P.Picture pPicture;

        public PictureScNew(SlideMasterSc slideMaster, P.Picture pPicture) : base()
        {
            this.slideMaster = slideMaster;
            this.pPicture = pPicture;
        }

        public override IPlaceholder Placeholder => throw new NotImplementedException();
    }

    
}