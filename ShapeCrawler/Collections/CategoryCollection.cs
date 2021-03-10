using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Shared;
using ShapeCrawler.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collection of chart categories.
    /// </summary>
    public class CategoryCollection : LibraryCollection<Category>
    {
        #region Constructors

        internal CategoryCollection(List<Category> categoryList)
        {
            CollectionItems = categoryList;
        }

        #endregion Constructors

        internal static CategoryCollection Create(OpenXmlElement firstChartSeries, ChartType chartType)
        {
            if (chartType == ChartType.BubbleChart || chartType == ChartType.ScatterChart)
            {
                return null;
            }

            var categoryList = new List<Category>();

            //  Get category data from the first series.
            //  Actually, it can be any series since all chart series contain the same categories.
            //  <c:cat>
            //      <c:strRef>
            //          <c:f>Sheet1!$A$2:$A$3</c:f>
            //          <c:strCache>
            //              <c:ptCount val="2"/>
            //              <c:pt idx="0">
            //                  <c:v>Category 1</c:v>
            //              </c:pt>
            //              <c:pt idx="1">
            //                  <c:v>Category 2</c:v>
            //              </c:pt>
            //          </c:strCache>
            //      </c:strRef>
            //  </c:cat>
            C.CategoryAxisData cCatAxisData = firstChartSeries.GetFirstChild<C.CategoryAxisData>();

            C.MultiLevelStringReference cMultiLvlStringRef = cCatAxisData.MultiLevelStringReference;
            if (cMultiLvlStringRef != null) // is it chart with multi-level category?
            {
                categoryList = GetMultiCategories(cMultiLvlStringRef);
            }
            else
            {
                C.NumberReference cNumReference = cCatAxisData.NumberReference;
                C.StringReference cStrReference = cCatAxisData.StringReference;
                IEnumerable<C.NumericValue> cachedValues = cNumReference != null // C.NumericValue (<c:v>) can store string value
                    ? cNumReference.NumberingCache.Descendants<C.NumericValue>()
                    : cStrReference.StringCache.Descendants<C.NumericValue>();

                int xCellIdx = 0;
                var xCells = new ResettableLazy<List<X.Cell>>(ChartReferencesParser.GetXCellsByFormula())
                foreach (C.NumericValue cachedValue in cachedValues)
                {
                    categoryList.Add(new Category(xCells, xCellIdx, cachedValue));
                }
            }

            return new CategoryCollection(categoryList);
        }

        #region Private Methods

        private static List<Category> GetMultiCategories(C.MultiLevelStringReference multiLvlStrRef)
        {
            var parents = new List<KeyValuePair<uint, Category>>();
            IEnumerable<C.Level> levels = multiLvlStrRef.MultiLevelStringCache.Elements<C.Level>().Reverse();
            foreach (C.Level lvl in levels)
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
                    foreach (C.StringPoint pt in ptElements)
                    {
                        var index = pt.Index;
                        var catName = pt.NumericValue.InnerText;
                        var category = new Category(catName);
                        nextParents.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }

                parents = nextParents;
            }

            return parents.Select(kvp => kvp.Value).ToList(parents.Count);
        }

        #endregion Private Methods
    }
}