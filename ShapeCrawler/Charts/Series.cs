using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart series.
    /// </summary>
    public class Series // TODO: convert to internal
    {
        private readonly Lazy<string> name;
        private readonly OpenXmlElement seriesXmlElement;

        internal Series(SCChart parentChart, OpenXmlElement seriesXmlElement)
        {
            this.ParentChart = parentChart;
            this.seriesXmlElement = seriesXmlElement;
            this.name = new Lazy<string>(this.GetNameOrDefault);
        }

        internal SCChart ParentChart { get; }

        /// <summary>
        ///     Gets chart type.
        /// </summary>
        public ChartType Type => this.ParentChart.Type;

        public IChartPointCollection Points =>
            ChartPointCollection.Create(this.ParentChart, this.seriesXmlElement);

        public bool HasName => this.name.Value != null;

        public string Name
        {
            get
            {
                if (this.name.Value == null)
                {
                    throw new NotSupportedException(ExceptionMessages.SeriesHasNotName);
                }

                return this.name.Value;
            }
        }

        private string GetNameOrDefault()
        {
            C.StringReference cStringReference = seriesXmlElement.GetFirstChild<C.SeriesText>()?.StringReference;
            if (cStringReference == null)
            {
                return null;
            }

            return ChartReferencesParser.GetSingleString(cStringReference, ParentChart);
        }
    }
}