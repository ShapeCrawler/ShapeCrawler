namespace ShapeCrawler.Charts
{
    public interface IChartPoint
    {
        public double Value { get; }
    }

    internal class ChartPoint : IChartPoint
    {
        public ChartPoint(double value)
        {
            this.Value = value;
        }

        public double Value { get; }
    }
}