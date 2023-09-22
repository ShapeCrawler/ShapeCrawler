using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Excel;
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
    private readonly ChartPart sdkChartPart;

    internal Categories(ChartPart sdkChartPart, OpenXmlElement firstChartSeries)
    {
        this.firstChartSeries = firstChartSeries;
        this.sdkChartPart = sdkChartPart;
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
        var cCatAxisData = (C.CategoryAxisData)firstChartSeries.First(x => x is C.CategoryAxisData);

        var cMultiLvlStringRef = cCatAxisData.MultiLevelStringReference;
        if (cMultiLvlStringRef != null)
        {
            categoryList = MultiCategories(cMultiLvlStringRef);
        }
        else
        {
            C.Formula cFormula;
            List<C.NumericValue> cachedValues; // C.NumericValue (<c:v>) can store string value
            var cNumReference = cCatAxisData.NumberReference;
            var cStrReference = cCatAxisData.StringReference!;
            if (cNumReference is not null)
            {
                cFormula = cNumReference.Formula!;
                cachedValues = cNumReference.NumberingCache!.Descendants<C.NumericValue>().ToList();
            }
            else
            {
                cFormula = cStrReference.Formula!;
                cachedValues = cStrReference.StringCache!.Descendants<C.NumericValue>().ToList();
            }

            var normalizedFormula = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty); // eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
            var sheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
            var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+").Value; // eg: Sheet1!A2:A5 -> A2:A5
            var addresses = new ExcelCellsRange(cellsRange).Addresses();
            for (var i = 0; i < addresses.Count; i++)
            {
                var category = new SheetCategory(this.sdkChartPart, sheetName, addresses[i], cachedValues[i]);
                categoryList.Add(category);
            }
        }

        return categoryList;
    }

    private static List<ICategory> MultiCategories(C.MultiLevelStringReference cMultiLevelStrRef)
    {
        var indexToCategory = new List<KeyValuePair<uint, ICategory>>();
        var topDownLevels = cMultiLevelStrRef.MultiLevelStringCache!.Elements<C.Level>().Reverse();
        foreach (C.Level cLevel in topDownLevels)
        {
            var cStringPoints = cLevel.Elements<C.StringPoint>();
            var nextIndexToCategory = new List<KeyValuePair<uint, ICategory>>();
            if (indexToCategory.Any())
            {
                List<KeyValuePair<uint, ICategory>> descOrderedMains =
                    indexToCategory.OrderByDescending(kvp => kvp.Key).ToList();
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue!;
                    KeyValuePair<uint, ICategory> parent = descOrderedMains.First(kvp => kvp.Key <= index);
                    var category = new MultiCategory(parent.Value, cachedCatName);
                    nextIndexToCategory.Add(new KeyValuePair<uint, ICategory>(index, category));
                }
            }
            else
            {
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue;
                    var category = new Category(cachedCatName!);
                    nextIndexToCategory.Add(new KeyValuePair<uint, ICategory>(index, category));
                }
            }

            indexToCategory = nextIndexToCategory;
        }

        return indexToCategory.Select(kvp => kvp.Value).ToList();
    }
}