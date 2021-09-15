using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Audio;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of a slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
    {
        private readonly P.ShapeTree shapeTree;
        private readonly SCSlide slide;

        #region Constructors

        internal ShapeCollection(List<IShape> shapes)
        {
            this.CollectionItems = shapes;
        }

        internal ShapeCollection(List<IShape> shapes, P.ShapeTree shapeTree, SCSlide slide)
        {
            this.slide = slide;
            this.CollectionItems = shapes;
            this.shapeTree = shapeTree;
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
            var pictureHandler = new PictureHandler(shapeContextBuilder);
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

            return new ShapeCollection(shapes, shapeTree, slide);
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

        public IAudioShape AddNewAudio(int x, int y, Stream audioStream)
        {
            MediaDataPart mediaDataPart1 = this.slide.ParentPresentation.PresentationDocument.CreateMediaDataPart("audio/mpeg", "mp3");
            audioStream.Position = 0;
            mediaDataPart1.FeedData(audioStream);
            ImagePart imagePart1 = this.slide.SlidePart.AddNewPart<ImagePart>("image/png", "rId44");
            var data = File.OpenRead(
                @"c:\Documents\ShapeCrawler\Issues\SC-159_Add API to add audio content on slide\image1.png");
            imagePart1.FeedData(data);

            AudioReferenceRelationship audioRr = this.slide.SlidePart.AddAudioReferenceRelationship(mediaDataPart1);
            MediaReferenceRelationship mediaRr = this.slide.SlidePart.AddMediaReferenceRelationship(mediaDataPart1);

            data.Close();

            P.Picture picture1 = new ();

            P.NonVisualPictureProperties nonVisualPictureProperties1 = new ();

            P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new () { Id = (UInt32Value)2U, Name = "test" };
            A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new ();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ();
            A.AudioFromFile audioFromFile1 = new A.AudioFromFile() { Link = audioRr.Id };

            P.ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ();

            P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new () { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

            P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
            media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(media1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(audioFromFile1);
            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

            P.BlipFill blipFill1 = new ();
            A.Blip blip1 = new A.Blip() { Embed = "rId44" };

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            P.ShapeProperties shapeProperties1 = new ();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 5791200L, Y = 3124200L };
            A.Extents extents2 = new A.Extents() { Cx = 609600L, Cy = 609600L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            this.shapeTree.Append(picture1);

            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)3972997422U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            //P.Timing timing1 = new P.Timing();

            //P.TimeNodeList timeNodeList1 = new P.TimeNodeList();

            //P.ParallelTimeNode parallelTimeNode1 = new P.ParallelTimeNode();

            //P.CommonTimeNode commonTimeNode1 = new P.CommonTimeNode() { Id = (UInt32Value)1U, Duration = "indefinite", Restart = P.TimeNodeRestartValues.Never, NodeType = P.TimeNodeValues.TmingRoot };

            //P.ChildTimeNodeList childTimeNodeList1 = new P.ChildTimeNodeList();

            //P.SequenceTimeNode sequenceTimeNode1 = new P.SequenceTimeNode() { Concurrent = true, NextAction = P.NextActionValues.Seek };

            //P.CommonTimeNode commonTimeNode2 = new P.CommonTimeNode() { Id = (UInt32Value)2U, Duration = "indefinite", NodeType = P.TimeNodeValues.MainSequence };

            //P.ChildTimeNodeList childTimeNodeList2 = new P.ChildTimeNodeList();

            //P.ParallelTimeNode parallelTimeNode2 = new P.ParallelTimeNode();

            //P.CommonTimeNode commonTimeNode3 = new P.CommonTimeNode() { Id = (UInt32Value)3U, Fill = P.TimeNodeFillValues.Hold };

            //P.StartConditionList startConditionList1 = new P.StartConditionList();
            //P.Condition condition1 = new P.Condition() { Delay = "indefinite" };

            //startConditionList1.Append(condition1);

            //P.ChildTimeNodeList childTimeNodeList3 = new P.ChildTimeNodeList();

            //P.ParallelTimeNode parallelTimeNode3 = new P.ParallelTimeNode();

            //P.CommonTimeNode commonTimeNode4 = new P.CommonTimeNode() { Id = (UInt32Value)4U, Fill = P.TimeNodeFillValues.Hold };

            //P.StartConditionList startConditionList2 = new P.StartConditionList();
            //P.Condition condition2 = new P.Condition() { Delay = "0" };

            //startConditionList2.Append(condition2);

            //P.ChildTimeNodeList childTimeNodeList4 = new P.ChildTimeNodeList();

            //P.ParallelTimeNode parallelTimeNode4 = new P.ParallelTimeNode();

            //P.CommonTimeNode commonTimeNode5 = new P.CommonTimeNode() { Id = (UInt32Value)5U, PresetId = 1, PresetClass = P.TimeNodePresetClassValues.MediaCall, PresetSubtype = 0, Fill = P.TimeNodeFillValues.Hold, NodeType = P.TimeNodeValues.ClickEffect };

            //P.StartConditionList startConditionList3 = new P.StartConditionList();
            //P.Condition condition3 = new P.Condition() { Delay = "0" };

            //startConditionList3.Append(condition3);

            //P.ChildTimeNodeList childTimeNodeList5 = new P.ChildTimeNodeList();

            //Command command1 = new Command() { Type = P.CommandValues.Call, CommandName = "playFrom(0.0)" };

            //P.CommonBehavior commonBehavior1 = new P.CommonBehavior();
            //P.CommonTimeNode commonTimeNode6 = new P.CommonTimeNode() { Id = (UInt32Value)6U, Duration = "120163", Fill = P.TimeNodeFillValues.Hold };

            //P.TargetElement targetElement1 = new P.TargetElement();
            //P.ShapeTarget shapeTarget1 = new P.ShapeTarget() { ShapeId = "2" };

            //targetElement1.Append(shapeTarget1);

            //commonBehavior1.Append(commonTimeNode6);
            //commonBehavior1.Append(targetElement1);

            //command1.Append(commonBehavior1);

            //childTimeNodeList5.Append(command1);

            //commonTimeNode5.Append(startConditionList3);
            //commonTimeNode5.Append(childTimeNodeList5);

            //parallelTimeNode4.Append(commonTimeNode5);

            //childTimeNodeList4.Append(parallelTimeNode4);

            //commonTimeNode4.Append(startConditionList2);
            //commonTimeNode4.Append(childTimeNodeList4);

            //parallelTimeNode3.Append(commonTimeNode4);

            //childTimeNodeList3.Append(parallelTimeNode3);

            //commonTimeNode3.Append(startConditionList1);
            //commonTimeNode3.Append(childTimeNodeList3);

            //parallelTimeNode2.Append(commonTimeNode3);

            //childTimeNodeList2.Append(parallelTimeNode2);

            //commonTimeNode2.Append(childTimeNodeList2);

            //P.PreviousConditionList previousConditionList1 = new P.PreviousConditionList();

            //P.Condition condition4 = new P.Condition() { Event = P.TriggerEventValues.OnPrevious, Delay = "0" };

            //P.TargetElement targetElement2 = new P.TargetElement();
            //P.SlideTarget slideTarget1 = new P.SlideTarget();

            //targetElement2.Append(slideTarget1);

            //condition4.Append(targetElement2);

            //previousConditionList1.Append(condition4);

            //P.NextConditionList nextConditionList1 = new P.NextConditionList();

            //P.Condition condition5 = new P.Condition() { Event = P.TriggerEventValues.OnNext, Delay = "0" };

            //P.TargetElement targetElement3 = new P.TargetElement();
            //P.SlideTarget slideTarget2 = new P.SlideTarget();

            //targetElement3.Append(slideTarget2);

            //condition5.Append(targetElement3);

            //nextConditionList1.Append(condition5);

            //sequenceTimeNode1.Append(commonTimeNode2);
            //sequenceTimeNode1.Append(previousConditionList1);
            //sequenceTimeNode1.Append(nextConditionList1);

            //P.Audio audio1 = new ();

            //P.CommonMediaNode commonMediaNode1 = new P.CommonMediaNode() { Volume = 80000 };

            //P.CommonTimeNode commonTimeNode7 = new P.CommonTimeNode() { Id = (UInt32Value)7U, Fill = P.TimeNodeFillValues.Hold, Display = false };

            //P.StartConditionList startConditionList4 = new P.StartConditionList();
            //P.Condition condition6 = new P.Condition() { Delay = "indefinite" };

            //startConditionList4.Append(condition6);

            //P.EndConditionList endConditionList1 = new P.EndConditionList();

            //P.Condition condition7 = new P.Condition() { Event = P.TriggerEventValues.OnStopAudio, Delay = "0" };

            //P.TargetElement targetElement4 = new P.TargetElement();
            //P.SlideTarget slideTarget3 = new P.SlideTarget();

            //targetElement4.Append(slideTarget3);

            //condition7.Append(targetElement4);

            //endConditionList1.Append(condition7);

            //commonTimeNode7.Append(startConditionList4);
            //commonTimeNode7.Append(endConditionList1);

            //P.TargetElement targetElement5 = new P.TargetElement();
            //P.ShapeTarget shapeTarget2 = new P.ShapeTarget() { ShapeId = "2" };

            //targetElement5.Append(shapeTarget2);

            //commonMediaNode1.Append(commonTimeNode7);
            //commonMediaNode1.Append(targetElement5);

            //audio1.Append(commonMediaNode1);

            //childTimeNodeList1.Append(sequenceTimeNode1);
            //childTimeNodeList1.Append(audio1);

            //commonTimeNode1.Append(childTimeNodeList1);

            //parallelTimeNode1.Append(commonTimeNode1);


            return null;
        }

        internal Shape GetShapeByPPlaceholderShape(P.PlaceholderShape inpPPlaceholderShape)
        {
            IEnumerable<Shape> placeholderShapes = CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
            Shape mappedShape = placeholderShapes.FirstOrDefault(IsEqual);

            bool IsEqual(Shape collectionShape)
            {
                Placeholder placeholder = (Placeholder) collectionShape.Placeholder;
                P.PlaceholderShape colPPlaceholderShape = placeholder.SdkPPlaceholderShape;

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