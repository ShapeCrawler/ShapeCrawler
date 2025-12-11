using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class SlideMasterTests : SCTest
{
    [Test]
    [Presentation("new")]
    [Presentation("023.pptx")]
    public void SlideNumber_Font_Color_Setter(IPresentation pres)
    {
        // Arrange
        var slideMaster = pres.SlideMasters[0];
        var green = new Color("00FF00");

        // Act
        slideMaster.SlideNumber!.Font.Color = green;

        // Assert
        Assert.That(slideMaster.SlideNumber.Font.Color.Hex, Is.EqualTo("00FF00"));
    }

    [Test]
    public void SlideNumber_Font_Size_Setter()
    {
        // Arrange
        var pres = new Presentation();
        var slideMaster = pres.SlideMasters[0];

        // Act
        pres.Footer.AddSlideNumber();
        slideMaster.SlideNumber!.Font.Size = 30;

        // Assert
        pres.Save();
        pres = SaveAndOpenPresentation(pres);
        slideMaster = pres.SlideMasters[0];
        slideMaster.SlideNumber!.Font.Size.Should().Be(30);
    }

    [Test]
    public void Shape_Width_and_Height_return_master_shape_width_and_height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var slideMaster = pres.SlideMasters[0];
        var shape = slideMaster.Shapes.First(sp => sp.Id == 2);

        // Act & Assert
        shape.Width.Should().Be(828);
        shape.Height.Should().BeApproximately(104.37m, 0.01m);
    }

    [Test]
    public void Shape_X_and_Y_return_master_shape_x_and_y_coordinates()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var slideMaster = pres.SlideMasters[0];
        var shape = slideMaster.Shapes.First(sp => sp.Id == 2);

        // Act & Assert
        shape.X.Should().Be(66);
        shape.Y.Should().BeApproximately(28.75m, 0.01m);
    }

    [Test]
    public void SlideLayout_Name_returns_name_of_slide_layout()
    {
        // Arrange
        var pptx = TestAsset("autoshape-case011_save-as-png.pptx");
        var pres = new Presentation(pptx);
        var slideMaster = pres.SlideMasters[0];

        // Act
        var layoutName = slideMaster.SlideLayouts[0].Name;

        // Assert
        layoutName.Should().Be("Title Slide");
    }

    [Test]
    public void AutoShapePlaceholderType_ReturnsPlaceholderType()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var slideMaster = pres.SlideMasters[0];
        var masterautoshapecase1 = slideMaster.Shapes.First(sp => sp.Id == 2);
        var masterautoshapecase2 = slideMaster.Shapes.First(sp => sp.Id == 8);
        var masterautoshapecase3 = slideMaster.Shapes.First(sp => sp.Id == 7);

        // Act
        PlaceholderType? shapePlaceholderTypeCase1 = masterautoshapecase1.PlaceholderType;

        // Assert
        shapePlaceholderTypeCase1.Should().Be(PlaceholderType.Title);
        masterautoshapecase2.PlaceholderType.Should().BeNull();
        masterautoshapecase3.PlaceholderType.Should().BeNull();
    }

    [Test]
    public void ShapeGeometryType_ReturnsShapesGeometryFormType()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        ISlideMaster slideMaster = pres.SlideMasters[0];
        IShape shapeCase1 = slideMaster.Shapes.First(sp => sp.Id == 2);
        IShape shapeCase2 = slideMaster.Shapes.First(sp => sp.Id == 8);

        // Act
        Geometry geometryTypeCase1 = shapeCase1.GeometryType;
        Geometry geometryTypeCase2 = shapeCase2.GeometryType;

        // Assert
        geometryTypeCase1.Should().Be(Geometry.Rectangle);
        geometryTypeCase2.Should().Be(Geometry.Custom);
    }

    [Test]
    public void AutoShapeTextBoxText_ReturnsText_WhenTheSlideMasterAutoShapesTextBoxIsNotEmpty()
    {
        // Arrange
        ISlideMaster slideMaster = new Presentation(TestAsset("001.pptx")).SlideMasters[0];
        IShape autoShape = (IShape)slideMaster.Shapes.First(sp => sp.Id == 8);

        // Act-Assert
        autoShape.TextBox.Text.Should().BeEquivalentTo("id8");
    }

    [Test]
    public void Number_returns_slide_master_order_number()
    {
        // Act & Assert
        new Presentation().SlideMaster(1).Number.Should().Be(1);
    }
    
    [Test]
    public void SlideLayout_Background_SolidFill_Color_Getter_returns_solid_color_of_the_slide_layout_background()
    {
        // Arrange
        var expectedColor = fixture.Color();
        var pres = new Presentation(p=>
        {
            p.SlideMaster(1).SlideLayout(1).Background.SolidFillColor(expectedColor);
        });
        var slideMaster = pres.SlideMaster(1);
        
        // Assert
        slideMaster.SlideLayout(1).Background.SolidFill.Color.Should().Be(expectedColor);
    }
    
    [Test]
    public void SlideLayout_Background_Picture_returns_slide_layout_background_picture_image()
    {
        // Arrange
        var expectedImage = fixture.Image();
        var pres = new Presentation(p=>
        {
            p.SlideMaster(1).SlideLayout(1).Background.Picture(expectedImage);
        });
        expectedImage.Position = 0;
        var expectedStream = new MemoryStream();
        expectedImage.CopyTo(expectedStream);
        var slideMaster = pres.SlideMaster(1);
        
        // Act
        var actualStream = slideMaster.SlideLayout(1).Background.Picture();

        // Assert
        actualStream.ToArray().Should().Equal(expectedStream.ToArray());
    }
}