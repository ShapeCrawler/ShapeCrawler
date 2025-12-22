using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Categories(ChartPart chartPart) : IReadOnlyList<ICategory>
{
    public int Count => this.CategoryList().Count;

    public ICategory this[int index] => this.CategoryList()[index];

    public IEnumerator<ICategory> GetEnumerator() => this.CategoryList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private static int ColumnIndex(string column)
    {
        const int asciiOffsetForA = 64;
        const int alphabetSize = 26;
        int result = 0;
        string col = column.ToUpper();
        for (int iChar = col.Length - 1; iChar >= 0; iChar--)
        {
            char columnPiece = col[iChar];
            int columnNumber = columnPiece - asciiOffsetForA;
            result += columnNumber * (int)Math.Pow(alphabetSize, col.Length - (iChar + 1));
        }

        return result;
    }

    private static string ColumnLetter(int columnNumber)
    {
        const int alphabetSize = 26;
        const int asciiOffsetForA = 65;
        var columnLetter = new StringBuilder();

        while (columnNumber > 0)
        {
            var modulo = (columnNumber - 1) % alphabetSize;
            columnLetter.Insert(0, (char)(asciiOffsetForA + modulo));
            columnNumber = (columnNumber - modulo) / alphabetSize;
        }

        return columnLetter.ToString();
    }

    private List<ICategory> MultiCategories(
        IEnumerable<C.Level> levels,
        string? sheetName,
        int startRow,
        int startColumnIndex)
    {
        var indexToCategory = new List<KeyValuePair<uint, ICategory>>();
        var topDownLevels = levels.Reverse().ToList();

        for (int i = 0; i < topDownLevels.Count; i++)
        {
            indexToCategory = this.ProcessLevel(topDownLevels[i], i, sheetName, startRow, startColumnIndex, indexToCategory);
        }

        return [.. indexToCategory.Select(kvp => kvp.Value)];
    }

    private List<KeyValuePair<uint, ICategory>> ProcessLevel(
        C.Level cLevel,
        int levelIndex,
        string? sheetName,
        int startRow,
        int startColumnIndex,
        List<KeyValuePair<uint, ICategory>> parentCategories)
    {
        string? addressPrefix = sheetName != null ? ColumnLetter(startColumnIndex + levelIndex) : null;
        var cStringPoints = cLevel.Elements<C.StringPoint>();
        var nextIndexToCategory = new List<KeyValuePair<uint, ICategory>>();
        var descOrderedMains = parentCategories.OrderByDescending(kvp => kvp.Key).ToList();

        foreach (C.StringPoint cStrPoint in cStringPoints)
        {
            var index = cStrPoint.Index!.Value;
            var cachedCatName = cStrPoint.NumericValue;
            if (cachedCatName == null)
            {
                continue;
            }

            string? address = addressPrefix != null ? $"{addressPrefix}{startRow + (int)index}" : null;
            ICategory category;
            if (descOrderedMains.Count != 0)
            {
                var parent = descOrderedMains.First(kvp => kvp.Key <= index).Value;
                category = new MultiCategory(chartPart, parent, cachedCatName, sheetName, address);
            }
            else
            {
                category = new Category(chartPart, cachedCatName, sheetName, address);
            }

            nextIndexToCategory.Add(new KeyValuePair<uint, ICategory>(index, category));
        }

        return nextIndexToCategory;
    }

    private List<ICategory> CategoryList()
    {
        var categoryList = new List<ICategory>();
        var cPlotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var firstSeries = cCharts.First().ChildElements
            .First(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
        var cCatAxisData = (C.CategoryAxisData)firstSeries.First(x => x is C.CategoryAxisData);

        var cMultiLvlStringRef = cCatAxisData.MultiLevelStringReference;
        if (cMultiLvlStringRef != null)
        {
            string? sheetName = null;
            int startRow = 0;
            int startColumnIndex = 0;

            if (cMultiLvlStringRef.Formula != null)
            {
                var formula = cMultiLvlStringRef.Formula.Text;
                var normalizedFormula = formula.Replace("'", string.Empty).Replace("$", string.Empty);
                sheetName = Regex.Match(normalizedFormula, @".+(?=\!)", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
                var range = Regex.Match(normalizedFormula, @"(?<=\!).+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
                var rangeStart = range.Split(':')[0];
                var match = Regex.Match(rangeStart, @"([A-Z]+)(\d+)", RegexOptions.None, TimeSpan.FromMilliseconds(1000));
                var startColumn = match.Groups[1].Value;
                startRow = int.Parse(match.Groups[2].Value);
                startColumnIndex = ColumnIndex(startColumn);
            }

            return this.MultiCategories(
                cMultiLvlStringRef.MultiLevelStringCache!.Elements<C.Level>(), 
                sheetName,
                startRow, 
                startColumnIndex);
        }

        var cStrLiteral = cCatAxisData.StringLiteral;
        if (cStrLiteral != null)
        {
            foreach (var pt in cStrLiteral.Elements<C.StringPoint>())
            {
                var category = new Category(chartPart, pt.NumericValue!, null, null);
                categoryList.Add(category);
            }

            return categoryList;
        }

        C.Formula cFormula;
        List<C.NumericValue> cachedValues;
        var cNumReference = cCatAxisData.NumberReference;
        var cStrReference = cCatAxisData.StringReference!;
        if (cNumReference is not null)
        {
            cFormula = cNumReference.Formula!;
            cachedValues = [.. cNumReference.NumberingCache!.Descendants<C.NumericValue>()];
        }
        else
        {
            cFormula = cStrReference.Formula!;
            cachedValues = [.. cStrReference.StringCache!.Descendants<C.NumericValue>()];
        }

        var normalizedFormulaRef = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty);
        var sheetNameRef = Regex
            .Match(normalizedFormulaRef, @".+(?=\!)", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
        var cellsRangeRef = Regex.Match(
            normalizedFormulaRef, 
            @"(?<=\!).+", 
            RegexOptions.None,
            TimeSpan.FromMilliseconds(1000)).Value;
        var addresses = new CellsRange(cellsRangeRef).Addresses();
        for (var i = 0; i < addresses.Count; i++)
        {
            var category = new SheetCategory(chartPart, sheetNameRef, addresses[i], cachedValues[i]);
            categoryList.Add(category);
        }

        return categoryList;
    }
}