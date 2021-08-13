using System;
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
    public class CategoryCollection : LibraryCollection<Category> // TODO: convert to internal
    {
        internal CategoryCollection(List<Category> categoryList)
        {
            this.CollectionItems = categoryList;
        }

        internal static CategoryCollection? Create(SCChart chart, OpenXmlElement firstChartSeries, ChartType chartType)
        {
            if (chartType is ChartType.BubbleChart or ChartType.ScatterChart)
            {
                // Bubble and Scatter charts do not have categories
                return null;
            }

            var categoryList = new List<Category>();

            // Get category data from the first series.
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
            C.CategoryAxisData cCatAxisData = (C.CategoryAxisData)firstChartSeries.First(x => x is C.CategoryAxisData);

            C.MultiLevelStringReference? cMultiLvlStringRef = cCatAxisData.MultiLevelStringReference;
            if (cMultiLvlStringRef != null)
            {
                categoryList = GetMultiCategories(cMultiLvlStringRef);
            }
            else
            {
                C.Formula cFormula;
                IEnumerable<C.NumericValue> cachedValues; // C.NumericValue (<c:v>) can store string value
                C.NumberReference? cNumReference = cCatAxisData.NumberReference;
                C.StringReference? cStrReference = cCatAxisData.StringReference;
                if (cNumReference is not null)
                {
                    cFormula = cNumReference.Formula;
                    cachedValues = cNumReference.NumberingCache.Descendants<C.NumericValue>();
                }
                else
                {
                    cFormula = cStrReference.Formula;
                    cachedValues = cStrReference.StringCache.Descendants<C.NumericValue>();
                }

                int catIndex = 0;

                ResettableLazy<List<X.Cell>> xCells = null;
                if (chart.ParentPresentation.Editable)
                {
                    xCells = new ResettableLazy<List<X.Cell>>(() =>
                        ChartReferencesParser.GetXCellsByFormula(cFormula, chart));
                }

                foreach (C.NumericValue cachedValue in cachedValues)
                {
                    categoryList.Add(new Category(xCells, catIndex++, cachedValue));
                }
            }

            return new CategoryCollection(categoryList);
        }

        #region Private Methods

        private static List<Category> GetMultiCategories(C.MultiLevelStringReference multiLevelStrRef)
        {
            var indexToCategory = new List<KeyValuePair<uint, Category>>();
            IEnumerable<C.Level> topDownLevels = multiLevelStrRef.MultiLevelStringCache.Elements<C.Level>().Reverse();
            foreach (C.Level cLevel in topDownLevels)
            {
                IEnumerable<C.StringPoint> cStrPoints = cLevel.Elements<C.StringPoint>();
                var nextIndexToCategory = new List<KeyValuePair<uint, Category>>();
                if (indexToCategory.Any())
                {
                    List<KeyValuePair<uint, Category>> descOrderedMains =
                        indexToCategory.OrderByDescending(kvp => kvp.Key).ToList();
                    foreach (C.StringPoint cStrPoint in cStrPoints)
                    {
                        uint index = cStrPoint.Index.Value;
                        C.NumericValue cachedCatName = cStrPoint.NumericValue;
                        KeyValuePair<uint, Category> parent = descOrderedMains.First(kvp => kvp.Key <= index);
                        Category category = new(null, -1, cachedCatName, parent.Value);
                        nextIndexToCategory.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }
                else
                {
                    foreach (C.StringPoint cStrPoint in cStrPoints)
                    {
                        uint index = cStrPoint.Index.Value;
                        C.NumericValue cachedCatName = cStrPoint.NumericValue;
                        var category = new Category(null, -1, cachedCatName);
                        nextIndexToCategory.Add(new KeyValuePair<uint, Category>(index, category));
                    }
                }

                indexToCategory = nextIndexToCategory;
            }

            return indexToCategory.Select(kvp => kvp.Value).ToList(indexToCategory.Count);
        }

        /// <summary>
        ///     Gets list of categories for multi-level category chart.
        /// </summary>
        /// <param name="multiLevelStrRef">
        ///     <c:cat>
        ///         <c:multiLvlStrRef>
        ///             <c:f>
        ///                 Лист1!$A$1:$B$6
        ///             </c:f>
        ///             <c:multiLvlStrCache>
        ///                 <c:ptCount val="6" />
        ///                 <c:lvl>
        ///                     <c:pt idx="0">
        ///                         <c:v>
        ///                             Dresses
        ///                         </c:v>
        ///                     </c:pt>
        ///                     <c:pt idx="1">
        ///                         <c:v>
        ///                             Tops
        ///                         </c:v>
        ///                     </c:pt>
        ///                     <c:pt idx="2">
        ///                         <c:v>
        ///                             Boots
        ///                         </c:v>
        ///                     </c:pt>
        ///                     <c:pt idx="3">
        ///                         <c:v>
        ///                             Flats
        ///                         </c:v>
        ///                     </c:pt>
        ///                 </c:lvl>
        ///                 <c:lvl>
        ///                     <c:pt idx="0">
        ///                         <c:v>
        ///                             Clothing
        ///                         </c:v>
        ///                     </c:pt>
        ///                     <c:pt idx="2">
        ///                         <c:v>
        ///                             Shoes
        ///                         </c:v>
        ///                     </c:pt>
        ///                 </c:lvl>
        ///             </c:multiLvlStrCache>
        ///         </c:multiLvlStrRef>
        ///     </c:cat>
        /// </param>
        private static List<Category> GetMultiCategoriesNew(C.MultiLevelStringReference multiLevelStrRef)
        {
            List<C.Level> cLevels = multiLevelStrRef.MultiLevelStringCache.Elements<C.Level>().ToList();
            List<Category> resultCatList = null;
            List<Category> prevCatList = null;
            for (int i = 0; i < cLevels.Count; i++)
            {
                IEnumerable<C.StringPoint> cStrPoints = cLevels[i].Elements<C.StringPoint>();
                var curCatList = new List<Category>();
                foreach (C.StringPoint cStrPoint in cStrPoints)
                {
                    uint catIndex = cStrPoint.Index.Value;
                    C.NumericValue catCachedName = cStrPoint.NumericValue;
                    var category = new Category(null, (int) catIndex, catCachedName);
                    curCatList.Add(category);
                }

                if (resultCatList == null)
                {
                    resultCatList = curCatList;
                    prevCatList = curCatList;
                }
                else
                {
                    prevCatList = curCatList;
                }
            }

            throw new NotImplementedException();
        }

        #endregion Private Methods
    }
}