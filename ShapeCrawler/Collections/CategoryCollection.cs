using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

        internal static CategoryCollection Create(SlideChart slideChart, OpenXmlElement firstChartSeries, ChartType chartType)
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
                C.Formula cFormula;
                IEnumerable<C.NumericValue> cachedValues; // C.NumericValue (<c:v>) can store string value
                C.NumberReference cNumReference = cCatAxisData.NumberReference;
                C.StringReference cStrReference = cCatAxisData.StringReference;
                if (cNumReference != null)
                {
                    cFormula = cNumReference.Formula;
                    cachedValues = cNumReference.NumberingCache.Descendants<C.NumericValue>();
                }
                else
                {
                    cFormula = cStrReference.Formula;
                    cachedValues = cStrReference.StringCache.Descendants<C.NumericValue>();
                }

                int xCellIdx = 0;
                var xCells = new ResettableLazy<List<X.Cell>>(() => ChartReferencesParser.GetXCellsByFormula(cFormula, slideChart));
                foreach (C.NumericValue cachedValue in cachedValues)
                {
                    categoryList.Add(new Category(xCells, xCellIdx++, cachedValue));
                }
            }

            return new CategoryCollection(categoryList);
        }

        #region Private Methods

        private static List<Category> GetMultiCategories(C.MultiLevelStringReference multiLvlStrRef) //TODO: optimize
        {
            var parents = new List<KeyValuePair<uint, Category>>();
            IEnumerable<C.Level> cLevels = multiLvlStrRef.MultiLevelStringCache.Elements<C.Level>().Reverse();
            foreach (C.Level lvl in cLevels)
            {
                IEnumerable<C.StringPoint> cStrPoints = lvl.Elements<C.StringPoint>();
                var nextParents = new List<KeyValuePair<uint, Category>>();
                if (parents.Any())
                {
                    List<KeyValuePair<uint, Category>> descParents = parents.OrderByDescending(kvp => kvp.Key).ToList();
                    foreach (C.StringPoint cStrPoint in cStrPoints)
                    {
                        uint index = cStrPoint.Index.Value;
                        C.NumericValue cachedCatName = cStrPoint.NumericValue;
                        KeyValuePair<uint, Category> parent = descParents.First(kvp => kvp.Key <= index);
                        Category category = new (null, -1, cachedCatName, parent.Value);
                        nextParents.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }
                else
                {
                    foreach (C.StringPoint pt in cStrPoints)
                    {
                        var index = pt.Index;
                        C.NumericValue cachedCatName = pt.NumericValue;
                        var category = new Category(null, -1, cachedCatName);
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