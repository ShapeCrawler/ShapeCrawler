using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using SlideDotNet.Comparer;
using SlideDotNet.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a collection of the chart category.
    /// </summary>
    public class CategoryCollection : LibraryCollection<Category>
    {
        #region Constructors

        /// <summary>
        /// Initializes a new collection of the chart categories.
        /// </summary>
        /// <param name="sdkFirstChartSeries">First series. Actually, it does not matter: all chart series contain the same categories.</param>
        public CategoryCollection(OpenXmlElement sdkFirstChartSeries)
        {
            Check.NotNull(sdkFirstChartSeries, nameof(sdkFirstChartSeries));

            var catAxData = sdkFirstChartSeries.GetFirstChild<C.CategoryAxisData>();
            var multiLvlStrRef = catAxData.MultiLevelStringReference;
            var numRef = catAxData.NumberReference;
            var strRef = catAxData.StringReference;

            if (multiLvlStrRef != null)
            {
                AddMultiCategories(multiLvlStrRef);
            }
            else
            {
                IEnumerable<C.NumericValue> sdkNumericValues;
                if (numRef != null)
                {
                    sdkNumericValues = numRef.NumberingCache.Descendants<C.NumericValue>();
                }
                else
                {
                    sdkNumericValues = strRef.StringCache.Descendants<C.NumericValue>();
                }
                CollectionItems = new List<Category>(sdkNumericValues.Count());
                foreach (var numericValue in sdkNumericValues)
                {
                    CollectionItems.Add(new Category(numericValue.InnerText));
                }
            }
        }

        #endregion Constructors

        #region Private Methods

        private void AddMultiCategories(C.MultiLevelStringReference multiLvlStrRef)
        {
            var parents = new SortedDictionary<uint, Category>(new DescComparer<uint>());
            var levels = multiLvlStrRef.MultiLevelStringCache.Elements<C.Level>().Reverse();
            foreach (var lvl in levels)
            {
                var ptElements = lvl.Elements<C.StringPoint>();
                var nextParents = new SortedDictionary<uint, Category>(new DescComparer<uint>());
                if (parents.Any())
                {
                    foreach (var pt in ptElements)
                    {
                        var index = pt.Index;
                        var parent = parents.First(kvp => kvp.Key <= index);
                        nextParents.Add(index, new Category(pt.NumericValue.InnerText, parent.Value));
                    }
                }
                else
                {
                    foreach (var pt in ptElements)
                    {
                        var index = pt.Index;
                        nextParents.Add(index, new Category(pt.NumericValue.InnerText));
                    }
                }

                parents = nextParents;
            }

            var ascParents = parents.OrderBy(kvp => kvp.Key);
            CollectionItems = new List<Category>(ascParents.Select(kvp => kvp.Value));
        }

        #endregion Private Methods
    }
}