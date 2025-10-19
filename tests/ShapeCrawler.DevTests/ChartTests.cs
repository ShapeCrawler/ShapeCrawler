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
        var scatterChart = pres.Slide(2).Shapes.GetById(5).Chart;

        // Act
        double xValue = scatterChart!.XAxis.Values[0];

        // Assert
        xValue.Should().Be(10);
    }
    
    [Test]
    public void Categories_is_null_When_the_chart_type_doesnt_have_categories()
    {
        // Arrange
        var pres = new Presentation(TestAsset("021.pptx"));
        var chart = pres.Slide(3).Shape(4).Chart;

        // Act & Assert
        chart.Categories.Should().BeNull();
    }

    [Test]
    public void TitleAndHasTitle_ReturnChartTitleStringAndFlagIndicatingWhetherChartHasATitle()
    {
        // Arrange
        var pres13 = new Presentation(TestAsset("013.pptx"));
        var pres19 = new Presentation(TestAsset("019.pptx"));
        var chartCase1 = new Presentation(TestAsset("018.pptx")).Slide(1).Shape(6).Chart;
        var chartCase2 = new Presentation(TestAsset("025_chart.pptx")).Slide(1).Shape(7).Chart;
        var chartCase3 = pres13.Slide(1).Shape(5).Chart;
        var chartCase5 = pres19.Slide(1).Shape(4).Chart;
        var chartCase7 = new Presentation(TestAsset("009_table.pptx")).Slide(3).Shape(7).Chart;
        var chartCase8 = new Presentation(TestAsset("009_table.pptx")).Slide(3).Shape(6).Chart;
        var chartCase9 = new Presentation(TestAsset("009_table.pptx")).Slide(5).Shape(6).Chart;
        var chartCase10 =new Presentation(TestAsset("009_table.pptx")).Slide(5).Shape(3).Chart;
        var chartCase11 = new Presentation(TestAsset("009_table.pptx")).Slide(5).Shape(5).Chart;
            
        // Act
        string charTitleCase1 = chartCase1.Title!.Text!;
        string charTitleCase2 = chartCase2.Title!.Text!;
        string charTitleCase3 = chartCase3.Title!.Text!;
        string charTitleCase5 = chartCase5.Title!.Text!;
        string charTitleCase7 = chartCase7.Title!.Text!;
        string charTitleCase8 = chartCase8.Title!.Text!;
        string charTitleCase9 = chartCase9.Title!.Text!;
        string charTitleCase10 = chartCase10.Title!.Text!;
        string charTitleCase11 = chartCase11.Title!.Text!;

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
    public void Title_Setter_updates_chart_title()
    {
        // Arrange
        var chart1 = new Presentation(TestAsset("018.pptx")).Slide(1).Shape(6).Chart;
        var chart2 = new Presentation(TestAsset("025_chart.pptx")).Slide(1).Shape(7).Chart;

        var newTitle = "To infinity and beyond!";

        // Act
        chart1.Title!.Text = newTitle;
        chart2.Title!.Text = null;

        // Assert
        chart1.Title.Text.Should().Be(newTitle);
        chart2.Title.Text.Should().BeNull();
    }
        
    [Test]
    public void SeriesCollection_Series_Points_returns_chart_point_collection()
    {
        // Arrange
        var pptxStream = TestAsset("005 chart.pptx");
        var presentation = new Presentation(pptxStream);
        var chart = (IChart) presentation.Slides[0].Shapes.First(shape => shape.Name == "chart").Chart;
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
        var chart = pres.Slides[0].Shapes.Shape(chartName).Chart;
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
        var chart1 = new Presentation(TestAsset("025_chart.pptx")).Slide(1).Shape(4).Chart;
        var chart2 = new Presentation(TestAsset("009_table.pptx")).Slide(3).Shape(7).Chart;

        // Act-Assert
        chart1.Categories[0].Name.Should().BeEquivalentTo("Dresses");
        chart2.Categories[0].Name.Should().BeEquivalentTo("Q1");
        chart2.Categories[1].Name.Should().BeEquivalentTo("Q2");
        chart2.Categories[2].Name.Should().BeEquivalentTo("Q3");
        chart2.Categories[3].Name.Should().BeEquivalentTo("Q4");
    }
        
    [Test]
    public void Category_Name_Getter_returns_category_name_for_chart_from_collection_of_Combination_chart()
    {
        // Arrange
        var comboChart = new Presentation(TestAsset("021.pptx")).Slide(1).Shape(4).Chart;

        // Act-Assert
        comboChart.Categories[0].Name.Should().BeEquivalentTo("2015");
    }

    [Test]
    public void CategoryName_GetterReturnsChartCategoryName_OfMultiCategoryChart()
    {
        // Arrange
        var chart = new Presentation(TestAsset("025_chart.pptx")).Slide(1).Shape(4).Chart;

        // Act-Assert
        chart.Categories[0].MainCategory.Name.Should().BeEquivalentTo("Clothing");
    }

    [Test]
    public void Category_Name_Setter_updates_category_name_in_non_multi_category_pie_chart()
    {
        // Arrange
        var pres = new Presentation(TestAsset("025_chart.pptx"));
        var mStream = new MemoryStream();
        var pieChart = pres.Slides[0].Shapes.GetById(7).Chart;

        // Act
        pieChart.Categories[0].Name = "Category 1_new";

        // Assert
        pieChart.Categories[0].Name.Should().Be("Category 1_new");
        pres.Save(mStream);
        pres = new Presentation(mStream);
        pieChart = pres.Slides[0].Shape(7).Chart;
        pieChart.Categories[0].Name.Should().Be("Category 1_new");
    }

    [Test, Ignore("ClosedXML dependency must be removed")]
    public void Category_Name_Setter_updates_value_of_Excel_cell()
    {
        // Arrange
        var pres = new Presentation(TestAsset("025_chart.pptx"));
        var lineChart = pres.Slides[3].Shape(13).Chart;
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
        IChart chart = (IChart)new Presentation(TestAsset("021.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3).Chart;
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
        var chart = new Presentation(TestAsset("025_chart.pptx")).Slide(1).Shape(5).Chart;

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
        IChart chartCase1 = (IChart)new Presentation(TestAsset("021.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3).Chart;
        IChart chartCase2 = (IChart)new Presentation(TestAsset("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4).Chart;
        IChart chartCase3 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 5).Chart;
        IChart chartCase4 = (IChart)new Presentation(TestAsset("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7).Chart;

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
        var pres = new Presentation(TestAsset("018.pptx"));
        var shape = pres.Slide(1).Shape(6);

        // Act-Assert
        shape.GeometryType.Should().Be(Geometry.Rectangle);
    }

    [Test]
    public void XAxis_Minimum()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001 bar chart.pptx"));
        var barChart = pres.Slide(1).Shape("Bar Chart 1").Chart;
        
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
        var barChart = pres.Slides[0].Shape("Bar Chart 1").Chart;
        var mStream = new MemoryStream();
        
        // Act
        barChart.XAxis!.Minimum = 1;

        // Assert
        pres.Save(mStream);
        barChart = new Presentation(mStream).Slides[0].Shape("Bar Chart 1").Chart;
        barChart.XAxis!.Minimum.Should().Be(1);
        pres.Validate();
    }
    
    [Test]
    public void Axes_ValueAxis_Maximum_Setter()
    {
        // Arrange
        var pptx = TestAsset("001 bar chart.pptx");
        var pres = new Presentation(pptx);
        var barChart = pres.Slides[0].Shapes.Shape("Bar Chart 1").Chart;
        
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
        var barChart = pres.Slide(1).Shape("Bar Chart 1").Chart;
        
        // Act & Assert
        barChart.XAxis!.Maximum.Should().Be(6);
    }
    
    [Test]
    [SlideShape("013.pptx", slideNumber:1, shapeId: 5, expectedResult: 3)]
    [SlideShape("009_table.pptx", slideNumber:3, shapeId: 7, expectedResult: 1)]
    public void SeriesCollection_Count_returns_number_of_series(IShape shape, int expectedSeriesCount)
    {
        // Act
        var chart = shape.Chart;
        int seriesCount = chart.SeriesCollection.Count;

        // Assert
        seriesCount.Should().Be(expectedSeriesCount);
    }

    [Test]
    public void Title_FontColor_Setter_update_chart_title_color()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.PieChart("Pie Chart 1");
            });
        });
        const string green = "00ff00";
        var chart = pres.Slide(1).Shape("Pie Chart 1").Chart!;
        chart.Title!.Text = "Sales Chart";
        
        // Act
        chart.Title.FontColor = green;
        
        // Assert
        chart.Title.FontColor.Should().Be(green);
        chart.Title.Text.Should().Be("Sales Chart");
    }
    
    [Test]
    public void Title_FontSize_Setter_update_chart_title_font_size()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.PieChart("Pie Chart 1");
            });
        });
        var title= pres.Slide(1).Shape("Pie Chart 1").Chart!.Title!;
        
        // Act
        title.FontSize = 14;
        
        // Assert
        title.FontSize.Should().Be(14);
    }
    
    [Test]
    public void Title_Text_Getter_returns_default_pie_chart_title()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.PieChart("Pie Chart 1");
            });
        });
        
        // Act-Assert
        pres.Slide(1).Shape("Pie Chart 1").Chart!.Title!.Text.Should().Be("Series 1");
    }
}