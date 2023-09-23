using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration
// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal sealed class Categories : IReadOnlyList<ICategory>
{
    private readonly IEnumerable<OpenXmlElement> cCharts;
    private readonly ChartPart sdkChartPart;

    internal Categories(ChartPart sdkChartPart, IEnumerable<OpenXmlElement> cCharts)
    {
        this.cCharts = cCharts;
        this.sdkChartPart = sdkChartPart;
    }

    public int Count => this.CategoryList().Count;
    public ICategory this[int index] => this.CategoryList()[index];
    public IEnumerator<ICategory> GetEnumerator() => this.CategoryList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<ICategory> CategoryList()
    {
        var categoryList = new List<ICategory>();
        var firstSeries = this.cCharts.First().ChildElements
            .First(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        var cCatAxisData = (C.CategoryAxisData)firstSeries.First(x => x is C.CategoryAxisData);

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