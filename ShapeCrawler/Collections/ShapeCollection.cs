using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

        internal static ShapeCollection CreateForSlideMaster(SlideMasterSc slideMaster)
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

        internal Shape GetShapeByPPlaceholderShape(P.PlaceholderShape pPlaceholderShapeParam)
        {
            Shape mappedShape = CollectionItems.OfType<Shape>().First(collectionShape =>
            {
                P.PlaceholderShape pPlaceholderShape = ((Placeholder) collectionShape.Placeholder).PPlaceholderShape;
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

    internal class MasterPicture : MasterShape, IShape
    {
        public MasterPicture(SlideMasterSc slideMaster, P.Picture pPicture) : base(slideMaster, pPicture)
        {
            throw new NotImplementedException();
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public GeometryType GeometryType { get; }
    }

    internal class MasterTable : MasterShape, IShape
    {
        public MasterTable(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame) : base(slideMaster, pGraphicFrame)
        {
            throw new NotImplementedException();
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public GeometryType GeometryType { get; }
    }

    internal class MasterChart : MasterShape, IShape
    {
        public MasterChart(SlideMasterSc slideMaster, P.GraphicFrame pGraphicFrame)
        :base(slideMaster, pGraphicFrame)
        {
            throw new NotImplementedException();
        }

        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public int Id { get; }
        public string Name { get; }
        public bool Hidden { get; }
        public GeometryType GeometryType { get; }
    }
}