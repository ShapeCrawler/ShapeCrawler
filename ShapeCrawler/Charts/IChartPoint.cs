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
        private readonly double? cachedValue;
        private readonly string address;
        private readonly SCChart parentChart;

        internal ChartPoint(SCChart parentChart, string address)
        {
            this.parentChart = parentChart;
            this.address = address;
        }

        internal ChartPoint(SCChart parentChart, string address, double cachedValue)
            : this(parentChart, address)
        {
            this.cachedValue = cachedValue;
        }

        public double Value
        {
            get => GetValue();
            set => this.UpdateValue(value);
        }

        private double GetValue()
        {
            if (this.cachedValue != null)
            {
                return this.cachedValue.Value;
            }

            
            return -1;
        }

        private void UpdateValue(double value)
        {
            throw new System.NotImplementedException();
        }
    }
}