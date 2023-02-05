using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
public class ChartTests : ShapeCrawlerTest
{
    [Fact]
    public void XValues_ReturnsParticularXAxisValue_ViaItsCollectionIndexer()
    {
        // Arrange
        var pptx = GetTestStream("024_chart.pptx");
        var pres = SCPresentation.Open(pptx);
        IChart chart = pres.Slides[1].Shapes.First(sp => sp.Id == 5) as IChart;

        // Act
        double xValue = chart.XValues[0];

        // Assert
        xValue.Should().Be(10);
        chart.HasXValues.Should().BeTrue();
    }

    [Fact]
    public void HasXValues()
    {
        // Arrange
        var pptx = GetTestStream("025_chart.pptx");
        var pres = SCPresentation.Open(pptx);
        ISlide slide1 = pres.Slides[0];
        ISlide slide2 = pres.Slides[1];
        IChart chart8 = slide1.Shapes.First(x => x.Id == 8) as IChart;
        IChart chart11 = slide2.Shapes.First(x => x.Id == 11) as IChart;

        // Act
        var chart8HasXValues = chart8.HasXValues;
        var chart11HasXValues = chart11.HasXValues;

        // Assert
        Assert.False(chart8HasXValues);
        Assert.False(chart11HasXValues);
    }

    [Fact]
    public void HasCategories_ReturnsFalse_WhenAChartHasNotCategories()
    {
        // Arrange
        IChart chart = (IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4);

        // Act
        bool hasChartCategories = chart.HasCategories;

        // Assert
        hasChartCategories.Should().BeFalse();
    }

    [Fact]
    public void TitleAndHasTitle_ReturnChartTitleStringAndFlagIndicatingWhetherChartHasATitle()
    {
        // Arrange
        var pres13 = SCPresentation.Open(GetTestStream("013.pptx"));
        var pres19 = SCPresentation.Open(GetTestStream("019.pptx"));
        IChart chartCase1 = (IChart)SCPresentation.Open(GetTestStream("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 6);
        IChart chartCase2 = (IChart)SCPresentation.Open(GetTestStream("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);
        IChart chartCase3 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 5);
        IChart chartCase4 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 4);
        IChart chartCase5 = (IChart)pres19.Slides[0].Shapes.First(sp => sp.Id == 4);
        IChart chartCase6 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 6);
        IChart chartCase7 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        IChart chartCase8 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 6);
        IChart chartCase9 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 6);
        IChart chartCase10 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 3);
        IChart chartCase11 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[4].Shapes.First(sp => sp.Id == 5);
            
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
        bool hasTitleCase4 = chartCase4.HasTitle;
        bool hasTitleCase6 = chartCase6.HasTitle;

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
        hasTitleCase4.Should().BeFalse();
        hasTitleCase6.Should().BeFalse();
    }
        
    [Theory]
    [MemberData(nameof(TestCasesSeriesCollectionCount))]
    public void SeriesCollection_Count_returns_number_of_series(IChart chart, int expectedSeriesCount)
    {
        // Act
        int seriesCount = chart.SeriesCollection.Count;

        // Assert
        Assert.Equal(expectedSeriesCount, seriesCount);
    }

    public static IEnumerable<object[]> TestCasesSeriesCollectionCount()
    {
        var pptxStream = GetTestStream("013.pptx");
        var presentation = SCPresentation.Open(pptxStream);
        IChart chart = (IChart) presentation.Slides[0].Shapes.First(sp => sp.Id == 5);
        yield return new object[] {chart, 3};

        pptxStream = GetTestStream("009_table.pptx");
        presentation = SCPresentation.Open(pptxStream);
        chart = (IChart) presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        yield return new object[] {chart, 1};
    }

    [Fact]
    public void SeriesCollection_Series_Points_returns_chart_point_collection()
    {
        // Arrange
        var pptxStream = GetTestStream("charts-case001.pptx");
        var presentation = SCPresentation.Open(pptxStream);
        var chart = (IChart) presentation.Slides[0].Shapes.First(shape => shape.Name == "chart");
        var series = chart.SeriesCollection[0]; 
            
        // Act
        var chartPoints = series.Points;
            
        // Assert
        chartPoints.Should().NotBeEmpty();
    }
            
    [Fact]
    public void CategoryName_GetterReturnsChartCategoryName()
    {
        // Arrange
        IBarChart chartCase1 = (IBarChart)SCPresentation.Open(GetTestStream("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);
        IPieChart chartCase3 = (IPieChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);

        // Act-Assert
        chartCase1.Categories[0].Name.Should().BeEquivalentTo("Dresses");
        chartCase3.Categories[0].Name.Should().BeEquivalentTo("Q1");
        chartCase3.Categories[1].Name.Should().BeEquivalentTo("Q2");
        chartCase3.Categories[2].Name.Should().BeEquivalentTo("Q3");
        chartCase3.Categories[3].Name.Should().BeEquivalentTo("Q4");
    }
        
    [Fact]
    public void Category_Name_Getter_returns_category_name_for_chart_from_collection_of_Combination_chart()
    {
        // Arrange
        var comboChart = (IComboChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);

        // Act-Assert
        comboChart.Categories[0].Name.Should().BeEquivalentTo("2015");
    }

    [Fact]
    public void CategoryName_GetterReturnsChartCategoryName_OfMultiCategoryChart()
    {
        // Arrange
        var chartCase1 = (IBarChart)SCPresentation.Open(GetTestStream("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 4);

        // Act-Assert
        chartCase1.Categories[0].MainCategory.Name.Should().BeEquivalentTo("Clothing");
    }

    [Fact]
    public void CategoryName_SetterChangesName_OfCategoryInNonMultiCategoryPieChart()
    {
        // Arrange
        var pres = SCPresentation.Open(GetTestStream("025_chart.pptx"));
        MemoryStream mStream = new();
        IPieChart pieChart4 = (IPieChart)pres.Slides[0].Shapes.First(sp => sp.Id == 7);
        const string newCategoryName = "Category 1_new";

        // Act
        pieChart4.Categories[0].Name = newCategoryName;

        // Assert
        pieChart4.Categories[0].Name.Should().Be(newCategoryName);
        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        pieChart4 = (IPieChart)pres.Slides[0].Shapes.First(sp => sp.Id == 7);
        pieChart4.Categories[0].Name.Should().Be(newCategoryName);
    }

    [Fact]
    public void Category_Name_Setter_updates_value_of_Excel_cell()
    {
        // Arrange
        var pres = SCPresentation.Open(GetTestStream("025_chart.pptx"));
        var lineChart = pres.Slides[3].Shapes.GetById<ILineChart>(13);
        const string newName = "Category 1_new";
        var category = lineChart.Categories[0]; 

        // Act
        category.Name = newName;

        // Assert
        var mStream = new MemoryStream(lineChart.WorkbookByteArray);
        var workbook = new XLWorkbook(mStream);
        var cellValue = workbook.Worksheets.First().Cell("A2").Value.ToString();
        cellValue.Should().BeEquivalentTo(newName);
    }

    [Fact(Skip = "On Hold")]
    public void CategoryName_SetterChangeName_OfSecondaryCategoryInMultiCategoryBarChart()
    {
        // Arrange
        Stream preStream = TestFiles.Presentations.pre025_byteArray.ToResizeableStream();
        IPresentation presentation = SCPresentation.Open(preStream);
        IBarChart barChart = (IBarChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 4);
        const string newCategoryName = "Clothing_new";

        // Act
        barChart.Categories[0].Name = newCategoryName;

        // Assert
        barChart.Categories[0].Name.Should().Be(newCategoryName);

        presentation.Save();
        presentation = SCPresentation.Open(preStream);
        barChart = (IBarChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 4);
        barChart.Categories[0].Name.Should().Be(newCategoryName);
    }

    [Fact]
    public void SeriesType_ReturnsChartTypeOfTheSeries()
    {
        // Arrange
        IChart chart = (IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        ISeries series2 = chart.SeriesCollection[1];
        ISeries series3 = chart.SeriesCollection[2];

        // Act
        SCChartType seriesChartType2 = series2.Type;
        SCChartType seriesChartType3 = series3.Type;

        // Assert
        seriesChartType2.Should().Be(SCChartType.BarChart);
        seriesChartType3.Should().Be(SCChartType.ScatterChart);
    }

    [Fact]
    public void Series_Name_returns_chart_series_name()
    {
        // Arrange
        IChart chart = (IChart)SCPresentation.Open(GetTestStream("025_chart.pptx")).Slides[0].Shapes.First(sp => sp.Id == 5);

        // Act
        string seriesNameCase1 = chart.SeriesCollection[0].Name;
        string seriesNameCase2 = chart.SeriesCollection[2].Name;

        // Assert
        seriesNameCase1.Should().BeEquivalentTo("Ряд 1");
        seriesNameCase2.Should().BeEquivalentTo("Ряд 3");
    }

    [Fact]
    public void Type_ReturnsChartType()
    {
        // Arrange
        var pres13 = SCPresentation.Open(GetTestStream("013.pptx"));
        IChart chartCase1 = (IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3);
        IChart chartCase2 = (IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4);
        IChart chartCase3 = (IChart)pres13.Slides[0].Shapes.First(sp => sp.Id == 5);
        IChart chartCase4 = (IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);

        // Act
        SCChartType chartTypeCase1 = chartCase1.Type;
        SCChartType chartTypeCase2 = chartCase2.Type;
        SCChartType chartTypeCase3 = chartCase3.Type;
        SCChartType chartTypeCase4 = chartCase4.Type;

        // Assert
        chartTypeCase1.Should().Be(SCChartType.BubbleChart);
        chartTypeCase2.Should().Be(SCChartType.ScatterChart);
        chartTypeCase3.Should().Be(SCChartType.Combination);
        chartTypeCase4.Should().Be(SCChartType.PieChart);
    }

    [Fact]
    public void GeometryType_Getter_returns_rectangle()
    {
        // Arrange
        IChart chart = (IChart)SCPresentation.Open(GetTestStream("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 6);

        // Act-Assert
        chart.GeometryType.Should().Be(SCGeometry.Rectangle);
    }
        
                
    [Fact]
    public void SDKSpreadsheetDocument_return_underlying_SpreadsheetDocument()
    {
        // Arrange
        var pptxStream = GetTestStream("charts-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var chart = pres.Slides[0].Shapes.GetByName<IChart>("Chart 1");
            
        // Act
        var spreadSheetDocument = chart.SDKSpreadsheetDocument;
            
        // Assert
        spreadSheetDocument.Should().NotBeNull();
    }
}