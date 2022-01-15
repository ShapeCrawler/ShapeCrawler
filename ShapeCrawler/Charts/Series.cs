using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Charts
{
    public interface ISeries
    {
        string Name { get; }

        /// <summary>
        ///     Gets chart type.
        /// </summary>
        ChartType Type { get; }

        IChartPointCollection Points { get; }
        
        bool HasName { get; }
    }

    /// <summary>
    ///     Represents a chart series.
    /// </summary>
    internal class Series : ISeries
    {
        private readonly Lazy<string> name;
        private readonly OpenXmlElement seriesXmlElement;

        public Series(SCChart parentChart, OpenXmlElement seriesXmlElement, ChartType seriesChartType)
        {
            this.ParentChart = parentChart;
            this.seriesXmlElement = seriesXmlElement;
            this.name = new Lazy<string>(this.GetNameOrDefault);
            this.Type = seriesChartType;
        }

        public SCChart ParentChart { get; }

        /// <summary>
        ///     Gets chart type.
        /// </summary>
        public ChartType Type { get; }

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