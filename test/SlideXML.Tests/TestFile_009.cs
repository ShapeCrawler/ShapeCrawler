using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NSubstitute;
using SlideXML.Enums;
using SlideXML.Exceptions;
using SlideXML.Extensions;
using SlideXML.Models;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;
using SlideXML.Services;
using SlideXML.Services.Placeholders;
using SlideXML.Tests.Helpers;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace SlideXML.Tests
{
    public class TestFile_009
    {
        [Fact]
        public void SlidesNumber_Test()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new Presentation(ms);

            // ACT
            var sldNumber = pre.Slides.Count();

            // CLOSE
            pre.Close();
            
            // ASSERT
            Assert.Equal(2, sldNumber);
        }

        /// <State>
        /// - there is a presentation with two slides. The first slide contains one element. The second slide includes two elements;
        /// - the first slide is removed;
        /// - the presentation is closed.
        /// </State>
        /// <ExpectedBahavior>
        /// There is the only second slide with its two elements.
        /// </ExpectedBahavior>
        [Fact]
        public void SlidesRemove_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._007_2_slides);
            var pre = new Presentation(ms);

            // ACT
            var slide1 = pre.Slides.First();
            pre.Slides.Remove(slide1);
            pre.Close();

            var pre2 = new Presentation(ms);
            var numSlides = pre2.Slides.Count();
            var numElements = pre2.Slides.Single().Shapes.Count;
            pre2.Close();
            ms.Dispose();

            // ASSERT
            Assert.Equal(1, numSlides);
            Assert.Equal(2, numElements);
        }

        [Fact]
        public void ShapeTextBody_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._008);

            // ACT
            var shapes = pre.Slides.Single().Shapes.OfType<SlideElement>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
           
            pre.Close();

            // ASSERT
            Assert.False(sh36.HasTextFrame);
            Assert.True(sh37.HasTextFrame);
            Assert.Equal("P1t1 P1t2", sh37.TextFrame.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextFrame.Paragraphs[1].Text);
        }

        [Fact]
        public void SlideElementsCount_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._003);

            // ACT
            var numberElements = pre.Slides.Single().Shapes.Count;
            pre.Close();

            // ASSERT
            Assert.Equal(5, numberElements);
        }

        [Fact]
        public void TextBox_Placeholder_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._006_1_slides);

            // ACT
            var shapePlaceholder = pre.Slides.Single().Shapes.Single();
            pre.Close();

            // ASSERT
            Assert.Equal(1524000, shapePlaceholder.X);
            Assert.Equal(1122363, shapePlaceholder.Y);
            Assert.Equal(9144000, shapePlaceholder.Width);
            Assert.Equal(2387600, shapePlaceholder.Height);
        }

        [Fact]
        public void GroupsElementPropertiesTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var slides = pre.Slides;
            var groupElement = pre.Slides[1].Shapes.Single(x => x.Type.Equals(ElementType.Group));
            var el3 = groupElement.Group.Shapes.Single(x => x.Id.Equals(5));
            pre.Close();

            // ASSERT
            Assert.Equal(1581846, el3.X);
            Assert.Equal(1181377, el3.Width);
            Assert.Equal(654096, el3.Height);
        }

        [Fact]
        public void SecondSlideElementsNumberTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var elNumber1 = pre.Slides[0].Shapes.Count;
            var elNumber2 = pre.Slides[1].Shapes.Count;
            pre.Close();

            // ASSERT
            Assert.Equal(6, elNumber1);
            Assert.Equal(5, elNumber2);
        }

        [Fact]
        public void SlideElementsDoNotThrowsExceptionTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var elements = pre.Slides[0].Shapes;

            pre.Close();
        }

        [Fact]
        public void PictureEx_BytesTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));

            // ACT
            var bytes = picEx.Picture.ImageEx.GetBytes().Result;

            // ASSERT
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void PictureEx_SetImageTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = picEx.Picture.ImageEx.GetBytes().Result.Length;

            // ACT
            picEx.Picture.ImageEx.SetImage(testImage2Stream);

            var sizeAfter = picEx.Picture.ImageEx.GetBytes().Result.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore,  sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_BytesTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var shapeEx = pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));

            // ACT
            var length = shapeEx.BackgroundImage.GetBytes().Result.Length;

            // ASSERT
            Assert.True(length > 0);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_SetImageTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var shapeEx = (SlideElement)pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = shapeEx.BackgroundImage.GetBytes().Result.Length;

            // ACT
            shapeEx.BackgroundImage.SetImage(testImage2Stream);

            var sizeAfter = shapeEx.BackgroundImage.GetBytes().Result.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_IsNullTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var shapeEx = (SlideElement)pre.Slides[1].Shapes.Single(e => e.Id.Equals(6));

            // ACT
            var bImage = shapeEx.BackgroundImage;

            // ASSERT
            Assert.Null(bImage);
        }

        [Fact]
        public void OleObjects_ParseTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var shapes = pre.Slides[1].Shapes;

            // ACT
            var oleNumbers = shapes.Count(e => e.Type.Equals(ElementType.OLEObject));
            var ole9 = shapes.Single(s => s.Id == 9);

            pre.Close();

            // ASSERT
            Assert.Equal(2, oleNumbers);
            Assert.Equal(699323, ole9.X);
            Assert.Equal(3463288, ole9.Y);
            Assert.Equal(485775, ole9.Width);
            Assert.Equal(373062, ole9.Height);
        }

        [Fact]
        public void OLEObject_NameTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var name = pre.Slides[1].Shapes.Single(e => e.Id.Equals(8)).Name;

            pre.Close();

            // ASSERT
            Assert.Equal("Object 2", name);
        }

        [Fact]
        public void SlideEx_Background_IsNullTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var bg = pre.Slides[1].BackgroundImage;

            // ASSERT
            Assert.Null(bg);
        }

        [Fact]
        public void SlideEx_Background_ChangeTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var bg = pre.Slides[0].BackgroundImage;
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = bg.GetBytes().Result.Length;

            // ACT
            bg.SetImage(testImage2Stream);

            var sizeAfter = bg.GetBytes().Result.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void NumberParagraphAndPortionTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var shape = (SlideElement)pre.Slides[2].Shapes.SingleOrDefault(e => e.Id.Equals(2));
            var paragraphs = shape.TextFrame.Paragraphs;

            // ACT
            var numParagraphs = paragraphs.Count;
            var portions = paragraphs[0].Portions;
            var numPortions = portions.Count;
            var por1Size = portions[0].FontHeight;
            var por2Size = portions[1].FontHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(1, numParagraphs);
            Assert.Equal(2, numPortions);
            Assert.Equal(1800, por1Size);
            Assert.Equal(2000, por2Size);
        }

        [Fact]
        public void SlideWidthAndHeightTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);

            // ACT
            var w = pre.SlideWidth;
            var y = pre.SlideHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(9144000, w);
            Assert.Equal(5143500, y);
        }

        [Fact]
        public void TextBox_Placeholder_FontHeight_Case1_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var elements = pre.Slides[3].Shapes;
            var tb2TitlePh = elements.Single(e => e.Id.Equals(2));
            var tb3SubTitlePh = elements.Single(e => e.Id.Equals(3));

            // ACT
            var fhTitle = tb2TitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;
            var text2 = tb2TitlePh.TextFrame.Text;
            var fhSubTitle = tb3SubTitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(4400, fhTitle);
            Assert.Equal(3200, fhSubTitle);
            Assert.Equal("Title text", text2);
        }

        [Fact]
        public void TablesPropertiesTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var elements = pre.Slides[2].Shapes;
            var tblEx = elements.Single(e => e.Id.Equals(3));
            var firstRow = tblEx.Table.Rows.First();

            // ACT
            var numRows = tblEx.Table.Rows.Count;
            var numCells = firstRow.Cells.Count;
            var numParagraphs = firstRow.Cells.First().TextBody.Paragraphs.Count;
            var cellTxt = firstRow.Cells.First().TextBody.Text;
            var prText = firstRow.Cells.First().TextBody.Paragraphs.First().Text;
            var portionTxt = firstRow.Cells.First().TextBody.Paragraphs.First().Portions.Single().Text;

            pre.Close();

            // ASSERT
            Assert.Equal(3, numRows);
            Assert.Equal(3, numCells);
            Assert.Equal(2, numParagraphs);
            Assert.Equal("0:0_p1_lvl1\r\n0:0_p2_lvl2", cellTxt);
            Assert.Equal("0:0_p1_lvl1", prText);
            Assert.Equal("0:0_p1_lvl1", portionTxt);
        }

        [Fact]
        public void ChartPropertiesTest()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._009);
            var sld3Elements = pre.Slides[2].Shapes;
            var sld5Elements = pre.Slides[4].Shapes;
            var chartEx6 = sld3Elements.Single(e => e.Id.Equals(6));
            var chartEx7 = sld3Elements.Single(e => e.Id.Equals(7));
            var sld5Chart6 = sld5Elements.Single(e => e.Id.Equals(6));
            var sld5Chart3 = sld5Elements.Single(e => e.Id.Equals(3));
            var sld5Chart5 = sld5Elements.Single(e => e.Id.Equals(5));

            // ACT
            var chart7Title = chartEx7.Chart.Title;
            var chart6Title = chartEx6.Chart.Title;
            var chart7Type = chartEx7.Chart.Type;
            var sld5Chart6Title = sld5Chart6.Chart.Title;
            var sld5Chart3Title = sld5Chart3.Chart.Title;
            var sld5Chart5Title = sld5Chart5.Chart.Title;
            var hasTextFrame = sld5Chart5.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.Equal("Sales", chart7Title);
            Assert.Equal("Sales2", chart6Title);
            Assert.Equal("Sales3", sld5Chart6Title);
            Assert.Equal("Sales4", sld5Chart3Title);
            Assert.Equal("Sales5", sld5Chart5Title);
            Assert.Equal(ChartType.PieChart, chart7Type);
            Assert.False(hasTextFrame);
        }

        [Fact]
        public void CreateShape_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubXmlShape = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(s => s.GetId() == 36);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubXmlShape,
                ElementType = ElementType.AutoShape
            };
            var creator = new ElementFactory(sldPart);
            var mockPreSetting = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSetting);

            // CLEAN
            doc.Close();

            // ASSERT
            Assert.Equal(ElementType.AutoShape, element.Type);
            Assert.Equal(3291840, element.X);
            Assert.Equal(274320, element.Y);
            Assert.Equal(1143000, element.Width);
            Assert.Equal(1143000, element.Height);
        }

        [Fact]
        public void DateTimePlaceholder_HasTextFrame_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextBody = sp3.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.False(hasTextBody);
        }

        [Fact]
        public void DateTimePlaceholder_Text_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._011_dt);
            var dt = pre.Slides[0].Shapes.Single(s => s.Id == 54275);

            // ACT
            var text = dt.TextFrame.Text;
            var hasText = dt.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.True(hasText);
            Assert.NotNull(text);
        }


        [Fact]
        public void TitlePlaceholder_TextAndFont_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._012_title_placeholder);
            var title = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var text = title.TextFrame.Text;
            var fh = title.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.Equal(2000, fh);
        }

        [Fact]
        public void TitlePlaceholder_FontHeight_Test()
        {
            // ARRANGE
            var pre010 = new Presentation(Properties.Resources._010);
            var pre010TextBox = pre010.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var fh = pre010TextBox.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre010.Close();

            // ASSERT
            Assert.Equal(1539, fh);
        }

        [Fact]
        public void ElementsNumber()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            var pre = new Presentation(ms);
            var allElements = pre.Slides.First().Shapes;

            // ACT
            var elementsNumber = allElements.Count;

            // CLOSE
            pre.Close();

            // ASSERT
            Assert.Equal(3, elementsNumber);
        }

        [Fact]
        public void Constructor_Test()
        {
            // ACT
            var exception = Assert.ThrowsAsync<TypeException>(() => throw new TypeException()).Result;

            // ASSERT
            Assert.Equal(101, exception.ErrorCode);
        }

        [Fact]
        public void Add_AddedOneItem_SlidesNumberIsOne()
        {
            // ARRANGE
            var xmlDoc = DocHelper.Open(Properties.Resources._001);
            var slides = new SlideCollection(xmlDoc);
            Substitute.For<IPlaceholderService>();
            var treeParser = new GroupShapeTypeParser();
            var bgImgFactory = new BackgroundImageFactory();
            var mockPreSettings = Substitute.For<IPreSettings>();
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            new ElementFactory(sldPart);

            var newSlide = new Slide(sldPart, 1, treeParser, bgImgFactory, mockPreSettings);

            // ACT
            slides.Add(newSlide);

            // CLEAN
            xmlDoc.Dispose();

            // ASSERT
            Assert.Single(slides);
        }

        /// <State>
        /// - there is presentation with two slides;
        /// - first slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation contains single slide with 1 number.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test1()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._007_2_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Equal(1, slides.Single().Number);

            // CLEAN
            pre.Close();
        }

        /// <State>
        /// - there is a single slide presentation;
        /// - slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation is empty.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test2()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._006_1_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Empty(slides);

            // CLEAN
            pre.Close();
        }

        [Fact]
        public void Shape_Text_Test2()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._012_title_placeholder);
            var autoShape = pre.Slides[0].Shapes.Single(s => s.Id == 3);

            // ACT
            var text = autoShape.TextFrame.Text;

            pre.Close();

            // ASSERT
            Assert.Equal("P1 P2", text);
        }

        [Fact]
        public void Shape_Text_Tests()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._011_dt);
            var autoShape = pre.Slides[0].Shapes.Single(s => s.Id == 2);
            var grShape = pre.Slides[0].Shapes.Single(s => s.Id == 4);

            // ACT
            var text = autoShape.TextFrame.Text;
            var hasTextFrame = grShape.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.False(hasTextFrame);
        }

        [Fact]
        public void IdHiddenIsPlaceholder_Test()
        {
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.SlideParts.Single();
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(ge => ge.GetId() == 6);

            // ACT
            var shapeBuilder = new SlideElement.Builder(new BackgroundImageFactory(), new GroupShapeTypeParser(), sldPart);
            var chartShape = shapeBuilder.BuildChart(stubGrFrame);

            // CLOSE
            ms.Dispose();
            doc.Dispose();

            // ASSERT
            Assert.Equal(6, chartShape.Id);
            Assert.False(chartShape.Hidden);
            Assert.False(chartShape.IsPlaceholder);
        }

        [Fact]
        public void Hidden_Test()
        {
            var ms = new MemoryStream(Properties.Resources._004);
            var pre = new Presentation(ms);

            // ACT
            var allElements = pre.Slides.Single().Shapes;
            var shapeHiddenValue = allElements[0].Hidden;
            var tableHiddenValue = allElements[1].Hidden;

            // CLOSE
            pre.Close();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }

        [Fact]
        public void Get_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);
            var sldLtPart = sldPart.SlideLayoutPart;
            var phService = new PlaceholderService(sldLtPart);

            // ACT
            var type = phService.TryGet(spId3).Type;

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, type);
        }

        [Fact]
        public void GetPlaceholderType_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);

            // ACT
            var phXml = PlaceholderService.GetPlaceholderXML(spId3);

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, phXml.PlaceholderType);
        }

        [Fact]
        public void CreateCandidates_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var parser = new GroupShapeTypeParser();
            var shapeTree = doc.PresentationPart.SlideParts.First().Slide.CommonSlideData.ShapeTree;

            // ACT
            var candidates = parser.CreateCandidates(shapeTree);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.AutoShape)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Picture)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Table)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Chart)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Group)));
        }

        [Fact]
        public void Build_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var groupShape = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GroupShape>().Single(x => x.GetId() == 2);
            var elFactory = new ElementFactory(sldPart);
            var builder = new SlideElement.Builder(new BackgroundImageFactory(), new GroupShapeTypeParser(), sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var groupEx = builder.BuildGroup(elFactory, groupShape, mockPreSettings);

            // CLEAN
            doc.Close();

            // ASSERT
            Assert.Equal(2, groupEx.Group.Shapes.Count);
            Assert.Equal(7547759, groupEx.X);
            Assert.Equal(2372475, groupEx.Y);
            Assert.Equal(1143000, groupEx.Width);
            Assert.Equal(1044645, groupEx.Height);
        }


        [Fact]
        public void CreatePicture_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubXmlPic = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.Picture>().Single();
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubXmlPic,
                ElementType = ElementType.Picture
            };
            var creator = new ElementFactory(sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Close();

            // ASSERT
            Assert.Equal(ElementType.Picture, element.Type);
            Assert.Equal(4663440, element.X);
            Assert.Equal(1005840, element.Y);
            Assert.Equal(2315880, element.Width);
            Assert.Equal(2315880, element.Height);
        }

        [Fact]
        public void CreateTable_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(e => e.GetId() == 38);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubGrFrame,
                ElementType = ElementType.Table
            };
            var creator = new ElementFactory(sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Close();

            // ASSERT
            Assert.Equal(ElementType.Table, element.Type);
            Assert.Equal(453240, element.X);
            Assert.Equal(3417120, element.Y);
            Assert.Equal(5075640, element.Width);
            Assert.Equal(1439640, element.Height);
        }

        [Fact]
        public void CreateChart_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(x => x.GetId() == 4);
            var stubEc = new ElementCandidate
            {
                CompositeElement = stubGrFrame,
                ElementType = ElementType.Chart
            };
            var creator = new ElementFactory(sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var element = creator.CreateShape(stubEc, mockPreSettings);

            // CLEAN
            doc.Close();

            // ASSERT
            Assert.Equal(ElementType.Chart, element.Type);
            Assert.Equal(453241, element.X);
            Assert.Equal(752401, element.Y);
            Assert.Equal(2672732, element.Width);
            Assert.Equal(1819349, element.Height);
        }
    }
}
