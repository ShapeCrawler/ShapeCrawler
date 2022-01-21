using System;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Exceptions;

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
        private double? cachedValue;
        private readonly string address;
        private readonly SCChart parentChart;
        private readonly string sheetName;

        internal ChartPoint(SCChart parentChart, string sheetName, string address)
        {
            this.parentChart = parentChart;
            this.sheetName = sheetName;
            this.address = address;
        }

        internal ChartPoint(SCChart parentChart, string sheetName, string address, double cachedValue)
            : this(parentChart, sheetName, address)
        {
            this.cachedValue = cachedValue;
        }

        public double Value
        {
            get => this.GetValue();
            set
            {
                try
                {
                    this.UpdateValue(value);
                }
                catch (Exception e)
                {
                    var logFile = Path.Combine(Path.GetTempPath(), "shapecrawler.log");
                    var messageBuilder = new StringBuilder();
                    messageBuilder.AppendLine($"Chart type:\t{this.parentChart.Type.ToString()}");
                    messageBuilder.AppendLine(e.ToString());
                    File.WriteAllText(logFile, messageBuilder.ToString());
                    throw new ShapeCrawlerException("An error occured while property updating. This should not happen, please report this as an issue on GitHub (https://github.com/ShapeCrawler/ShapeCrawler/issues).");
                }
            }
        }

        private double GetValue()
        {
            if (this.cachedValue != null)
            {
                return this.cachedValue.Value;
            }

            var xCell = this.parentChart.ChartWorkbook.GetXCell(this.sheetName, this.address);
            this.cachedValue = xCell.InnerText.Length == 0 ? 0 : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);

            return this.cachedValue.Value;
        }

        private void UpdateValue(double value)
        {
            throw new System.NotImplementedException("Inner Exception");
        }
    }
}