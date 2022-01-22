using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit
{
    public class ChartPointTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture fixture;

        public ChartPointTests(PresentationFixture fixture)
        {
            this.fixture = fixture;
        }
        
        [Fact]
        public void Value_Getter_returns_point_value_of_Bar_chart()
        {
            // Arrange
            IPresentation presentation = this.fixture.Pre021;
            var shapes1 = presentation.Slides[0].Shapes;
            var chart1 = (IChart) shapes1.First(x => x.Id == 3);
            ISeries chart6Series = ((IChart)this.fixture.Pre025.Slides[1].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];

            // Act
            double pointValue1 = chart1.SeriesCollection[1].Points[0].Value;
            double pointValue2 = chart6Series.Points[0].Value;

            // Assert
            Assert.Equal(56, pointValue1);
            Assert.Equal(72.66, pointValue2);
        }
        
        [Fact]
        public void Value_Getter_returns_point_value_of_Scatter_chart()
        {
            // Arrange
            IPresentation presentation = this.fixture.Pre021;
            var shapes1 = presentation.Slides[0].Shapes;
            var chart1 = (IChart) shapes1.First(x => x.Id == 3);
            
            // Act
            double scatterChartPointValue = chart1.SeriesCollection[2].Points[0].Value;
            
            // Assert
            Assert.Equal(44, scatterChartPointValue);
        }
        
        [Fact]
        public void Value_Getter_returns_point_value_of_Line_chart()
        {
            // Arrange
            var chart2 = this.GetShape<IChart>("021.pptx", 2, 4);
            var point = chart2.SeriesCollection[1].Points[0];

            // Act
            double lineChartPointValue = point.Value;

            // Assert
            Assert.Equal(17.35, lineChartPointValue);
        }
        
        [Fact]
        public void Value_Getter_returns_chart_point()
        {
            // Arrange
            ISeries seriesCase1 = ((IChart)this.fixture.Pre021.Slides[1].Shapes.First(sp => sp.Id == 3)).SeriesCollection[0];
            ISeries seriesCase2 = ((IChart)this.fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];
            ISeries seriesCase4 = ((IChart)this.fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7)).SeriesCollection[0];

            // Act
            double seriesPointValueCase1 = seriesCase1.Points[0].Value;
            double seriesPointValueCase2 = seriesCase2.Points[0].Value;
            double seriesPointValueCase4 = seriesCase4.Points[0].Value;
            double seriesPointValueCase5 = seriesCase4.Points[1].Value;

            // Assert
            seriesPointValueCase1.Should().Be(20.4);
            seriesPointValueCase2.Should().Be(2.4);
            seriesPointValueCase4.Should().Be(8.2);
            seriesPointValueCase5.Should().Be(3.2);
        }

        [Fact]
        public void Value_Setter_updates_chart_point()
        {
            // Arrange
            var chart = this.GetShape<IChart>("024_chart.pptx", 3, 5);
            var point = chart.SeriesCollection[0].Points[0];
            const int newValue = 6;

            // Act
            point.Value = newValue;

            // Assert
            point.Value.Should().Be(newValue);
            
            var stream = new MemoryStream();
            chart.ParentSlide.ParentPresentation.SaveAs(stream);
            chart = this.GetShape<IChart>(stream, 3, 5);
            var savedChartPoint = chart.SeriesCollection[0].Points[0];
            savedChartPoint.Value.Should().Be(newValue);

            var pointCellValue = this.GetCellValue<double>(chart.WorkbookByteArray, "B2");
            pointCellValue.Should().Be(newValue);
        }
    }
}