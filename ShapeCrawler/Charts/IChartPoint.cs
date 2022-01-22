using System;
using System.Globalization;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

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
        private readonly string address;
        private readonly SCChart parentChart;
        private readonly string sheetName;
        private readonly C.NumericValue? cNumericValue;

        internal ChartPoint(SCChart parentChart, string sheetName, string address)
        {
            this.parentChart = parentChart;
            this.sheetName = sheetName;
            this.address = address;
        }

        internal ChartPoint(SCChart parentChart, string sheetName, string address, C.NumericValue? cNumericValue)
            : this(parentChart, sheetName, address)
        {
            this.cNumericValue = cNumericValue;
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
            // From cache
            if (this.cNumericValue != null)
            {
                var cachedValue = double.Parse(this.cNumericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                return Math.Round(cachedValue, 2);
            }

            // From spreadsheet
            var xCell = this.parentChart.ChartWorkbook.GetXCell(this.sheetName, this.address);
            var sheetValue = xCell.InnerText.Length == 0 ? 0 : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);

            return sheetValue;
        }

        private void UpdateValue(double value)
        {
            // Try update cache
            if (this.cNumericValue != null)
            {
                this.cNumericValue.Text = value.ToString(CultureInfo.InvariantCulture);
            }

            // Update spreadsheet
            var xCell = this.parentChart.ChartWorkbook.GetXCell(this.sheetName, this.address);
            xCell.DataType = new EnumValue<X.CellValues>(X.CellValues.Number);
            xCell.CellValue = new X.CellValue(value);
        }
    }
}