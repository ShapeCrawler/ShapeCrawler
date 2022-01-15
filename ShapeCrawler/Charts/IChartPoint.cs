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
        public ChartPoint(double value)
        {
            this.Value = value;
        }

        public double Value { get; set; }
    }
}