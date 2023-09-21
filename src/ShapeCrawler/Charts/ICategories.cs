using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Charts;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal sealed class Categories : IReadOnlyCollection<ICategory>
{
    private readonly OpenXmlElement firstChartSeries;

    internal Categories(OpenXmlElement firstChartSeries)
    {
        this.firstChartSeries = firstChartSeries;
    }

    public int Count => this.CategoryList().Count;
    public ICategory this[int index] => this.CategoryList()[index];
    public IEnumerator<ICategory> GetEnumerator() => this.CategoryList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<ICategory> CategoryList()
    {
        var categoryList = new List<ICategory>();

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
        var cCatAxisData = (C.CategoryAxisData)firstChartSeries!.First(x => x is C.CategoryAxisData);

        var cMultiLvlStringRef = cCatAxisData.MultiLevelStringReference;
        if (cMultiLvlStringRef != null)
        {
            categoryList = GetMultiCategories(cMultiLvlStringRef);
        }
        else
        {
            C.Formula cFormula;
            IEnumerable<C.NumericValue> cachedValues; // C.NumericValue (<c:v>) can store string value
            var cNumReference = cCatAxisData.NumberReference;
            var cStrReference = cCatAxisData.StringReference!;
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
            
            foreach (C.NumericValue cachedValue in cachedValues)
            {
                categoryList.Add(new Category(cFormula, catIndex++, cachedValue));
            }
        }

        return categoryList;
    }

    private static List<ICategory> GetMultiCategories(C.MultiLevelStringReference multiLevelStrRef)
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