using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class GroupTests : SCTest
{
    [Test]
    public void Grouped_Shape_Y_Setter_raises_up_group_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape("Group 2");
        var groupedShape = groupShape.GroupedShapes.Shape("Shape 1");

        // Act
        groupedShape.Y = 307;

        // Assert
        groupedShape.Y.Should().Be(307);
        groupShape.Y.Should().Be(307, "because the moved grouped shape was on the up-hand side");
        groupShape.Height.Should().BeApproximately(91.87m, 0.01m);
    }

    [Test]
    public void Y_Setter_increases_the_height_of_the_group_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape("Group 2");
        var groupedShape = groupShape.GroupedShapes.Shape("Shape 2");

        // Act
        groupedShape.Y = 372;

        // Assert
        groupedShape.Y.Should().Be(372);
        groupShape.Height.Should().BeApproximately(90.08m, 0.01m);
    }

    [Test]
    public void X_Setter_moves_the_Left_hand_grouped_shape_to_Left()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape("Group 2");
        var groupedShape = groupShape.GroupedShapes.Shape("Shape 1");

        // Act
        groupedShape.X = 49m;

        // Assert
        groupedShape.X.Should().Be(49);
        groupShape.X.Should().Be(49, "because the moved grouped shape was on the left-hand side");
        groupShape.Width.Should().BeApproximately(89.18m, 0.01m);
    }

    [Test]
    public void X_Setter_moves_the_Right_hand_grouped_shape_to_Right()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var groupShape = pres.Slide(1).Shape("Group 2");
        var groupedShape = groupShape.GroupedShapes.Shape("Shape 1");
        var groupShapeX = groupShape.X;

        // Act
        groupedShape.X = 69m;

        // Assert
        groupedShape.X.Should().Be(69m);
        groupShape.X.Should().BeApproximately(groupShapeX, 0.01m,
            "because the X-coordinate of parent group shouldn't be changed when a grouped shape is moved to the right side");
        groupShape.Width.Should().BeApproximately(87.72m, 0.01m);
    }

    [Test]
    public void Width_returns_shape_width_in_points()
    {
        // Arrange
        var pres = new Presentation(TestAsset("006_1 slides.pptx"));
        var shapeCase1 = pres.Slide(1).Shapes.First(sp => sp.Id == 2);
        var pres2 = new Presentation(TestAsset("009_table.pptx"));
        var groupShape = pres2.Slide(2).Shape(7);
        var shapeCase2 = groupShape.GroupedShapes.GetById(5);
        var shapeCase3 = new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);

        // Act & Assert
        shapeCase1.Width.Should().BeApproximately(720m, 0.01m);
        shapeCase2.Width.Should().BeApproximately(93.02m, 0.01m);
        shapeCase3.Width.Should().BeApproximately(38.252m, 0.01m);
    }

    [Test]
    public void Height_returns_Grouped_Shape_height_in_pixels()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);
        var groupShape = pres.Slide(2).Shape("Group 1");
        var groupedShape = groupShape.GroupedShapes.Shape("Shape 2");

        // Act
        var height = groupedShape.Height;

        // Assert
        height.Should().BeApproximately(51.50m, 0.01m);
    }

    [Test]
    public void Shape_IsNotGroupShape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("006_1 slides.pptx"));
        var shape = pres.Slide(1).Shapes.GetById(3);

        // Act-Assert
        shape.GroupedShapes.Should().BeNull();
    }

    [Test]
    public void X_Getter_returns_x_coordinate_of_Grouped_shape_in_points()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shape = pres.Slide(2).Shape("Group 1").GroupedShapes.Shape("Shape 1");

        // Act
        decimal x = shape.X;

        // Assert
        x.Should().BeApproximately(39.94m, 0.01m);
    }

    [Test]
    public void Name_Setter_sets_grouped_shape_name()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var stream = new MemoryStream();
        var groupShape = pres.Slide(1).Shape("Group 2");

        // Act
        groupShape.Name = "New Group Name";

        // Assert
        pres.Save(stream);
        pres = new Presentation(stream);
        groupShape = pres.Slide(1).Shape("New Group Name");
        groupShape.Name.Should().Be("New Group Name");
        ValidatePresentation(pres);
    }

    [Test]
    public void GroupedShape_Width()
    {
        // Arrange
        var pres = new Presentation(TestAsset("077 grouped shape.pptx"));
        var groupShape = pres.Slide(1)
            .Shape("Group 19")
            .GroupedShape("Group 86")
            .GroupedShape("Group 1")
            .GroupedShape("Group 3")
            .GroupedShape("Rectangle 4");
        
        // Act & Assert
        groupShape.Width.Should().BeApproximately(290, 1m);
        groupShape.Height.Should().BeApproximately(36, 1m);
    }
    
    [Test]
    public void GroupedShape_X()
    {
        // Arrange
        var pres = new Presentation(TestAsset("077 grouped shape.pptx"));
        var groupedShape = pres.Slide(1)
            .Shape("Group 19")
            .GroupedShape("Group 86")
            .GroupedShape("Group 1")
            .GroupedShape("Group 3")
            .GroupedShape("Rectangle 4");
        
        // Act & Assert
        groupedShape.X.Should().BeApproximately(607, 1m);
    }
    
    [Test]
    public void GroupedShape_Y_Setter_returns_absolute_y_coordinate()
    {
        // Arrange
        using var pres = new Presentation(TestAsset("077 grouped shape.pptx"));
        var groupedShape = pres.Slide(1)
            .Shape("Group 19")
            .GroupedShape("Group 86")
            .GroupedShape("Group 1")
            .GroupedShape("Group 3")
            .GroupedShape("Rectangle 4");
        
        // Act & Assert
        groupedShape.Y.Should().BeApproximately(125, 1m);
    }
}