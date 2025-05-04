using System.Diagnostics.CodeAnalysis;
using ClosedXML.Excel;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.DevTests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
public class ChartTests : SCTest
{
    [Test]
    public void XValues_ReturnsParticularXAxisValue_ViaItsCollectionIndexer()
    {
        // Arrange
        var pres = new Presentation(TestAsset("024_chart.pptx"));
        var scatterChart = pres.Slide(2).Shapes.GetById<IChart>(5);

        // Act
        // double xValue = chart.XValues[0];
        double xValue = scatterChart.XAxis.Values[0];

        // Assert
        xValue.Should().Be(10);
    }
    
    [Test]
    public void Categories_is_null_When_the_chart_type_doesnt_have_categories()
    {
        // Arrange
        var pres = new Presentation(TestAsset("021.pptx"));
        var chart = pres.Slide(3).Chart(4);

        // Act & Assert
        chart.Categories.Should().BeNull();
    }

    [Test]
    public void TitleAndHasTitle_ReturnChartTitleStringAndFlagIndicatingWhetherChartHasATitle()
    {
        // Arrange
        var pres13 = new Presentation(TestAsset("013.pptx"));
        var pres19 = new Presentation(TestAsset("019.pptx"));
        IChart chartCase1 = (IChart)new Presentation(TestAsset("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 6);
        IChart chartCase2 = (IChart)new Presentation(TestAsset("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);
        IChart chartCase3 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 5);
        IChart chartCase4 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 4);
        IChart chartCase5 = (IChart)pres19.Slides[0].Shapes.First(sp => sp.Id == 4);
        IChart chartCase6 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 6);
        IChart chartCase7 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        IChart chartCase8 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 6);
        IChart chartCase9 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 6);
        IChart chartCase10 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 3);
        IChart chartCase11 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 5);
            
        // Act
        string charTitleCase1 = chartCase1.Title;
        string charTitleCase2 = chartCase2.Title;
        string charTitleCase3 = chartCase3.Title;
        string charTitleCase5 = chartCase5.Title;
        string charTitleCase7 = chartCase7.Title;
        string charTitleCase8 = chartCase8.Title;
        string charTitleCase9 = chartCase9.Title;
        string charTitleCase10 = chartCase10.Title;
        string charTitleCase11 = chartCase11.Title;

        // Assert
        charTitleCase1.Should().BeEquivalentTo("Test title");
        charTitleCase2.Should().BeEquivalentTo("Series 1_id7");
        charTitleCase3.Should().BeEquivalentTo("Title text");
        charTitleCase5.Should().BeEquivalentTo("Test title");
        charTitleCase7.Should().BeEquivalentTo("Sales");
        charTitleCase8.Should().BeEquivalentTo("Sales2");
        charTitleCase9.Should().BeEquivalentTo("Sales3");
        charTitleCase10.Should().BeEquivalentTo("Sales4");
        charTitleCase11.Should().BeEquivalentTo("Sales5");
    }
        
    [Test]
    public void SeriesCollection_Series_Points_returns_chart_point_collection()
    {
        // Arrange
        var pptxStream = TestAsset("005 chart.pptx");
        var presentation = new Presentation(pptxStream);
        var chart = (IChart) presentation.Slides[0].Shapes.First(shape => shape.Name == "chart");
        var series = chart.SeriesCollection[0]; 
            
        // Act
        var chartPoints = series.Points;
            
        // Assert
        chartPoints.Should().NotBeEmpty();
    }
    
    [TestCase("001 bar chart.pptx", "Bar Chart 1")]
    [TestCase("019.pptx", "Pie Chart 1")]
    public void SeriesCollection_RemoveAt_removes_series_by_index(string pptxFile, string chartName)
    {
        // Arrange
        var pptxStream = TestAsset(pptxFile);
        var pres = new Presentation(pptxStream);
        var chart = pres.Slides[0].Shapes.Shape<IChart>(chartName);
        var expectedSeriesCount = chart.SeriesCollection.Count - 1; 
            
        // Act
        chart.SeriesCollection.RemoveAt(0);

        // Assert
        chart.SeriesCollection.Count.Should().Be(expectedSeriesCount);
    }
    
    [Test]
    public void CategoryName_GetterReturnsChartCategoryName()
    {
        // Arrange
        IChart chartCase1 = (IChart)new Presentation(TestAsset("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);
        IChart chartCase3 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);

        // Act-Assert
        chartCase1.Categories[0].Name.Should().BeEquivalentTo("Dresses");
        chartCase3.Categories[0].Name.Should().BeEquivalentTo("Q1");
        chartCase3.Categories[1].Name.Should().BeEquivalentTo("Q2");
        chartCase3.Categories[2].Name.Should().BeEquivalentTo("Q3");
        chartCase3.Categories[3].Name.Should().BeEquivalentTo("Q4");
    }
        
    [Test]
    public void Category_Name_Getter_returns_category_name_for_chart_from_collection_of_Combination_chart()
    {
        // Arrange
        var comboChart = (IChart)new Presentation(TestAsset("021.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);

        // Act-Assert
        comboChart.Categories[0].Name.Should().BeEquivalentTo("2015");
    }

    [Test]
    public void CategoryName_GetterReturnsChartCategoryName_OfMultiCategoryChart()
    {
        // Arrange
        var chartCase1 = (IChart)new Presentation(TestAsset("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);

        // Act-Assert
        chartCase1.Categories[0].MainCategory.Name.Should().BeEquivalentTo("Clothing");
    }

    [Test]
    public void Category_Name_Setter_updates_category_name_in_non_multi_category_pie_chart()
    {
        // Arrange
        var pres = new Presentation(TestAsset("025_chart.pptx"));
        var mStream = new MemoryStream();
        var pieChart = pres.Slides[0].Shapes.GetById<IChart>(7);

        // Act
        pieChart.Categories[0].Name = "Category 1_new";

        // Assert
        pieChart.Categories[0].Name.Should().Be("Category 1_new");
        pres.Save(mStream);
        pres = new Presentation(mStream);
        pieChart = pres.Slides[0].Shapes.GetById<IChart>(7);
        pieChart.Categories[0].Name.Should().Be("Category 1_new");
    }

    [Test, Ignore("ClosedXML dependency must be removed")]
    public void Category_Name_Setter_updates_value_of_Excel_cell()
    {
        // Arrange
        var pres = new Presentation(TestAsset("025_chart.pptx"));
        var lineChart = pres.Slides[3].Shapes.GetById<IChart>(13);
        var category = lineChart.Categories[0]; 

        // Act
        category.Name = "Category 1_new";

        // Assert
        var mStream = new MemoryStream(lineChart.GetWorksheetByteArray());
        var workbook = new XLWorkbook(mStream);
        var cellValue = workbook.Worksheets.First().Cell("A2").Value.ToString();
        cellValue.Should().BeEquivalentTo("Category 1_new");
    }

    [Test, Ignore("On Hold")]
    public void CategoryName_SetterChangeName_OfSecondaryCategoryInMultiCategoryBarChart()
    {
        // Arrange
        var pptxStream = TestAsset("025_chart.pptx");
        var pres = new Presentation(pptxStream);
        var barChart = (IChart)pres.Slides[0].Shapes.First(sp => sp.Id == 4);
        const string newCategoryName = "Clothing_new";

        // Act
        barChart.Categories[0].Name = newCategoryName;

        // Assert
        barChart.Categories[0].Name.Should().Be(newCategoryName);

        pres.Save();
        pres = new Presentation(pptxStream);
        barChart = (IChart)pres.Slides[0].Shapes.First(sp => sp.Id == 4);
        barChart.Categories[0].Name.Should().Be(newCategoryName);
    }

    [Test]
    public void SeriesType_ReturnsChartTypeOfTheSeries()
    {
        // Arrange
        IChart chart = (IChart)new Presentation(TestAsset("021.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        ISeries series2 = chart.SeriesCollection[1];
        ISeries series3 = chart.SeriesCollection[2];

        // Act
        ChartType seriesChartType2 = series2.Type;
        ChartType seriesChartType3 = series3.Type;

        // Assert
        seriesChartType2.Should().Be(ChartType.BarChart);
        seriesChartType3.Should().Be(ChartType.ScatterChart);
    }

    [Test]
    public void Series_Name_returns_chart_series_name()
    {
        // Arrange
        IChart chart = (IChart)new Presentation(TestAsset("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 5);

        // Act
        string seriesNameCase1 = chart.SeriesCollection[0].Name;
        string seriesNameCase2 = chart.SeriesCollection[2].Name;

        // Assert
        seriesNameCase1.Should().BeEquivalentTo("Ряд 1");
        seriesNameCase2.Should().BeEquivalentTo("Ряд 3");
    }

    [Test]
    public void Type_ReturnsChartType()
    {
        // Arrange
        var pres13 = new Presentation(TestAsset("013.pptx"));
        IChart chartCase1 = (IChart)new Presentation(TestAsset("021.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);
        IChart chartCase2 = (IChart)new Presentation(TestAsset("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4);
        IChart chartCase3 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 5);
        IChart chartCase4 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);

        // Act
        ChartType chartTypeCase1 = chartCase1.Type;
        ChartType chartTypeCase2 = chartCase2.Type;
        ChartType chartTypeCase3 = chartCase3.Type;
        ChartType chartTypeCase4 = chartCase4.Type;

        // Assert
        chartTypeCase1.Should().Be(ChartType.BubbleChart);
        chartTypeCase2.Should().Be(ChartType.ScatterChart);
        chartTypeCase3.Should().Be(ChartType.Combination);
        chartTypeCase4.Should().Be(ChartType.PieChart);
    }

    [Test]
    public void GeometryType_Getter_returns_rectangle()
    {
        // Arrange
        IChart chart = (IChart)new Presentation(TestAsset("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 6);

        // Act-Assert
        chart.GeometryType.Should().Be(Geometry.Rectangle);
    }

    [Test]
    public void XAxis_Minimum()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001 bar chart.pptx"));
        var barChart = pres.Slide(1).Shapes.Shape<IChart>("Bar Chart 1");
        
        // Act
        var minimum = barChart.XAxis.Minimum;
        
        // Assert
        minimum.Should().Be(0);
    }
    
    [Test]
    public void Axes_ValueAxis_Minimum_Setter()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001 bar chart.pptx"));
        var barChart = pres.Slides[0].Shapes.Shape<IChart>("Bar Chart 1");
        var mStream = new MemoryStream();
        
        // Act
        barChart.XAxis!.Minimum = 1;

        // Assert
        pres.Save(mStream);
        barChart = new Presentation(mStream).Slides[0].Shapes.Shape<IChart>("Bar Chart 1");
        barChart.XAxis!.Minimum.Should().Be(1);
        pres.Validate();
    }
    
    [Test]
    public void Axes_ValueAxis_Maximum_Setter()
    {
        // Arrange
        var pptx = TestAsset("001 bar chart.pptx");
        var pres = new Presentation(pptx);
        var barChart = pres.Slides[0].Shapes.Shape<IChart>("Bar Chart 1");
        
        // Act
        barChart.XAxis!.Maximum = 7;

        // Assert
        barChart.XAxis.Maximum.Should().Be(7);
        pres.Validate();
    }
    
    [Test]
    public void XAxis_Maximum_Getter_returns_default_6()
    {
        // Arrange  
        var pres = new Presentation(TestAsset("001 bar chart.pptx"));
        var barChart = pres.Slide(1).Chart("Bar Chart 1");
        
        // Act & Assert
        barChart.XAxis!.Maximum.Should().Be(6);
    }
    
    [Test]
    [SlideShape("013.pptx", slideNumber:1, shapeId: 5, expectedResult: 3)]
    [SlideShape("009_table.pptx", slideNumber:3, shapeId: 7, expectedResult: 1)]
    public void SeriesCollection_Count_returns_number_of_series(IShape shape, int expectedSeriesCount)
    {
        // Act
        var chart = (IChart)shape;
        int seriesCount = chart.SeriesCollection.Count;

        // Assert
        seriesCount.Should().Be(expectedSeriesCount);
    }
}