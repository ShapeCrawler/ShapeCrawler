using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a chart categories.
    /// </summary>
    public class CategoryCollection : LibraryCollection<string>
    {
        #region Constructors

        /// <summary>
        /// Initializes a new collection of the chart categories.
        /// </summary>
        /// <param name="sdkFirstChartSeries">
        /// </param>
        public CategoryCollection(OpenXmlElement sdkFirstChartSeries)
        {
            Check.NotNull(sdkFirstChartSeries, nameof(sdkFirstChartSeries));

            IEnumerable<C.NumericValue> sdkNumericValues;
            var cCat = sdkFirstChartSeries.GetFirstChild<C.CategoryAxisData>();
            if (cCat.StringReference != null)
            {
                sdkNumericValues = cCat.StringReference.StringCache.Descendants<C.NumericValue>();
            }
            else
            {
                sdkNumericValues = cCat.NumberReference.NumberingCache.Descendants<C.NumericValue>();
            }
            CollectionItems = new List<string>(sdkNumericValues.Count());
            foreach (var numericValue in sdkNumericValues)
            {
                CollectionItems.Add(numericValue.InnerText);
            }
        }

        #endregion Constructors
    }
}