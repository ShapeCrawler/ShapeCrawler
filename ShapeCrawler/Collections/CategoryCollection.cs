using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
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
            var parents = new List<KeyValuePair<uint, Category>>();
            var levels = multiLvlStrRef.MultiLevelStringCache.Elements<C.Level>().Reverse();
            foreach (var lvl in levels)
            {
                var ptElements = lvl.Elements<C.StringPoint>();
                var nextParents = new List<KeyValuePair<uint, Category>>();
                if (parents.Any())
                {
                    var descParents = parents.OrderByDescending(kvp => kvp.Key).ToList();
                    foreach (var pt in ptElements)
                    {
                        var index = pt.Index;
                        var catName = pt.NumericValue.InnerText;
                        var parent = descParents.First(kvp => kvp.Key <= index);
                        var category = new Category(catName, parent.Value);
                        nextParents.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }
                else
                {
                    foreach (var pt in ptElements)
                    {
                        var index = pt.Index;
                        var catName = pt.NumericValue.InnerText;
                        var category = new Category(catName);
                        nextParents.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }

                parents = nextParents;
            }

            CollectionItems = parents.Select(kvp => kvp.Value).ToList(parents.Count);
        }

        #endregion Private Methods
    }
}