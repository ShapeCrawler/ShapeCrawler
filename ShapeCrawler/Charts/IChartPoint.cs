using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart point.
    /// </summary>
    public interface IChartPoint
    {
        /// <summary>
        ///     Gets or sets chart point value.
        /// </summary>
        public double Value { get; set; }
    }

    internal class ChartPoint : IChartPoint
    {
        private readonly double value;

        internal ChartPoint(double value)
        {
            this.value = value;
        }

        public double Value
        {
            get => this.value;
            set => this.UpdateValue(value);
        }

        private void UpdateValue(double value)
        {
            throw new System.NotImplementedException();
        }
    }
}