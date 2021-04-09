using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Tables;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of a slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
    {
        #region Constructors

        internal ShapeCollection(List<IShape> shapes)
        {
            CollectionItems = shapes;
        }

        #endregion Constructors

        internal static ShapeCollection CreateForSlide(SlidePart slidePart, SCSlide slide)
        {
            var phService = new PlaceholderService(slidePart.SlideLayoutPart);
            var transformFactory = new LocationParser(phService);
            var geometryFactory = new GeometryFactory(phService);
            var shapeContextBuilder = new ShapeContext.Builder(slidePart);

            var chartGrFrameHandler = new ChartGraphicFrameHandler();
            var tableGrFrameHandler = new TableGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var oleGrFrameHandler = new OleGraphicFrameHandler(shapeContextBuilder, transformFactory);
            var pShapeHandler = new AutoShapeCreator(shapeContextBuilder, transformFactory);
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

        internal static ShapeCollection CreateForSlideLayout(P.ShapeTree pShapeTree, SCSlideLayout slideLayout)
        {
            var shapeList = new List<IShape>();
            foreach (OpenXmlCompositeElement compositeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        shapeList.Add(new LayoutAutoShape(slideLayout, pShape));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutOLEObject(slideLayout, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutChart(slideLayout, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new LayoutTable(slideLayout, pGraphicFrame));
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
                    shapeList.Add(new LayoutPicture(slideLayout, pPicture));
                }
            }

            return new ShapeCollection(shapeList);
        }

        internal static ShapeCollection CreateForSlideMaster(SCSlideMaster slideMaster)
        {
            P.ShapeTree pShapeTree = slideMaster.PSlideMaster.CommonSlideData.ShapeTree;
            var shapeList = new List<IShape>();
            foreach (OpenXmlCompositeElement compositeElement in pShapeTree.OfType<OpenXmlCompositeElement>())
            {
                switch (compositeElement)
                {
                    case P.Shape pShape:
                        shapeList.Add(new MasterAutoShape(slideMaster, pShape));
                        continue;
                    case P.GraphicFrame pGraphicFrame:
                    {
                        A.GraphicData aGraphicData =
                            pGraphicFrame.GetFirstChild<A.Graphic>().GetFirstChild<A.GraphicData>();
                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterOLEObject(slideMaster, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterChart(slideMaster, pGraphicFrame));
                            continue;
                        }

                        if (aGraphicData.Uri.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                            StringComparison.Ordinal))
                        {
                            shapeList.Add(new MasterTable(slideMaster, pGraphicFrame));
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
                    shapeList.Add(new MasterPicture(slideMaster, pPicture));
                }
            }

            return new ShapeCollection(shapeList);
        }

        internal Shape GetShapeByPPlaceholderShape(P.PlaceholderShape inpPPlaceholderShape)
        {
            IEnumerable<Shape> placeholderShapes = CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
            Shape mappedShape = placeholderShapes.FirstOrDefault(IsEqual);

            bool IsEqual(Shape collectionShape)
            {
                Placeholder placeholder = (Placeholder) collectionShape.Placeholder;
                P.PlaceholderShape colPPlaceholderShape = placeholder.PPlaceholderShape;

                if (inpPPlaceholderShape.Index != null && colPPlaceholderShape.Index != null &&
                    inpPPlaceholderShape.Index == colPPlaceholderShape.Index)
                {
                    return true;
                }

                if (inpPPlaceholderShape.Type != null && colPPlaceholderShape.Type != null)
                {
                    if (inpPPlaceholderShape.Type == P.PlaceholderValues.Body &&
                        inpPPlaceholderShape.Index != null && colPPlaceholderShape.Index != null)
                    {
                        return inpPPlaceholderShape.Index == colPPlaceholderShape.Index;
                    }

                    return inpPPlaceholderShape.Type.Equals(colPPlaceholderShape.Type);
                }

                return false;
            }

            return mappedShape;
        }
    }
}