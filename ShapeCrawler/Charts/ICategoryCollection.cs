using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Collections;

/// <summary>
///     Represents a chart category collection.
/// </summary>
public interface ICategoryCollection : IEnumerable<ICategory>
{
    /// <summary>
    ///     Gets number of categories.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets category by specified index.
    /// </summary>
    ICategory this[int index] { get; }
}

internal sealed class CategoryCollection : LibraryCollection<ICategory>, ICategoryCollection
{
    private CategoryCollection(List<Category> categoryList)
        : base(categoryList)
    {
    }

    internal static CategoryCollection? Create(SCChart chart, OpenXmlElement? firstChartSeries, SCChartType chartType)
    {
        if (chartType is SCChartType.BubbleChart or SCChartType.ScatterChart)
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
        C.CategoryAxisData cCatAxisData = (C.CategoryAxisData)firstChartSeries!.First(x => x is C.CategoryAxisData);

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
            C.StringReference cStrReference = cCatAxisData.StringReference!;
            if (cNumReference is not null)
            {
                cFormula = cNumReference.Formula!;
                cachedValues = cNumReference.NumberingCache!.Descendants<C.NumericValue>();
            }
            else
            {
                cFormula = cStrReference.Formula!;
                cachedValues = cStrReference.StringCache!.Descendants<C.NumericValue>();
            }

            int catIndex = 0;
            ResettableLazy<List<X.Cell>> xCells;

            xCells = new ResettableLazy<List<X.Cell>>(() =>
                ChartReferencesParser.GetXCellsByFormula(cFormula, chart));
            foreach (C.NumericValue cachedValue in cachedValues)
            {
                categoryList.Add(new Category(xCells, catIndex++, cachedValue));
            }
        }

        return new CategoryCollection(categoryList);
    }

    private static List<Category> GetMultiCategories(C.MultiLevelStringReference multiLevelStrRef)
    {
        var indexToCategory = new List<KeyValuePair<uint, Category>>();
        IEnumerable<C.Level> topDownLevels = multiLevelStrRef.MultiLevelStringCache!.Elements<C.Level>().Reverse();
        foreach (C.Level cLevel in topDownLevels)
        {
            var cStringPoints = cLevel.Elements<C.StringPoint>();
            var nextIndexToCategory = new List<KeyValuePair<uint, Category>>();
            if (indexToCategory.Any())
            {
                List<KeyValuePair<uint, Category>> descOrderedMains =
                    indexToCategory.OrderByDescending(kvp => kvp.Key).ToList();
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue!;
                    KeyValuePair<uint, Category> parent = descOrderedMains.First(kvp => kvp.Key <= index);
                    var category = new Category(-1, cachedCatName, parent.Value);
                    nextIndexToCategory.Add(new KeyValuePair<uint, Category>(index, category));
                }
            }
            else
            {
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue;
                    var category = new Category(-1, cachedCatName!);
                    nextIndexToCategory.Add(new KeyValuePair<uint, Category>(index, category));
                }
            }

            indexToCategory = nextIndexToCategory;
        }

        return indexToCategory.Select(kvp => kvp.Value).ToList(indexToCategory.Count);
    }
}