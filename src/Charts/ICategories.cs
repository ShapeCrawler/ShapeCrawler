using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
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
        int retVal = 0;
        string col = column.ToUpper();
        for (int iChar = col.Length - 1; iChar >= 0; iChar--)
        {
            char colPiece = col[iChar];
            int colNum = colPiece - 64;
            retVal += colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
        }

        return retVal;
    }

    private static string ColumnLetter(int columnNumber)
    {
        string dividend = string.Empty;
        int modulo;

        while (columnNumber > 0)
        {
            modulo = (columnNumber - 1) % 26;
            dividend = Convert.ToChar(65 + modulo).ToString() + dividend;
            columnNumber = (int)((columnNumber - modulo) / 26);
        }

        return dividend;
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
            var cLevel = topDownLevels[i];
            
            string? addressPrefix = null;
            if (sheetName != null)
            {
                var currentColumnIndex = startColumnIndex + i;
                var currentColumnLetter = ColumnLetter(currentColumnIndex);
                addressPrefix = currentColumnLetter;
            }

            var cStringPoints = cLevel.Elements<C.StringPoint>();
            var nextIndexToCategory = new List<KeyValuePair<uint, ICategory>>();
            
            if (indexToCategory.Any())
            {
                List<KeyValuePair<uint, ICategory>> descOrderedMains =
                    [.. indexToCategory.OrderByDescending(kvp => kvp.Key)];
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue!;
                    KeyValuePair<uint, ICategory> parent = descOrderedMains.First(kvp => kvp.Key <= index);
                    
                    string? address = null;
                    if (addressPrefix != null)
                    {
                         var row = startRow + (int)index;
                         address = $"{addressPrefix}{row}";
                    }
                    
                    var category = new MultiCategory(chartPart, parent.Value, cachedCatName, sheetName, address);
                    nextIndexToCategory.Add(new KeyValuePair<uint, ICategory>(index, category));
                }
            }
            else
            {
                foreach (C.StringPoint cStrPoint in cStringPoints)
                {
                    var index = cStrPoint.Index!.Value;
                    var cachedCatName = cStrPoint.NumericValue;
                    
                    string? address = null;
                    if (addressPrefix != null)
                    {
                         var row = startRow + (int)index;
                         address = $"{addressPrefix}{row}";
                    }

                    if (cachedCatName == null)
                    {
                        continue;
                    }

                    var category = new Category(chartPart, cachedCatName, sheetName, address);
                    nextIndexToCategory.Add(new KeyValuePair<uint, ICategory>(index, category));
                }
            }

            indexToCategory = nextIndexToCategory;
        }

        return [.. indexToCategory.Select(kvp => kvp.Value)];
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
                 sheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value;
                 var range = Regex.Match(normalizedFormula, @"(?<=\!).+").Value;
                 var rangeStart = range.Split(':')[0];
                 var match = Regex.Match(rangeStart, @"([A-Z]+)(\d+)");
                 var startColumn = match.Groups[1].Value;
                 startRow = int.Parse(match.Groups[2].Value);
                 startColumnIndex = ColumnIndex(startColumn);
             }
             
             return this.MultiCategories(cMultiLvlStringRef.MultiLevelStringCache!.Elements<C.Level>(), sheetName, startRow, startColumnIndex);
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
            dividend = Convert.ToChar(65 + modulo) + dividend;
        var cellsRangeRef = Regex.Match(normalizedFormulaRef, @"(?<=\!).+", RegexOptions.None, TimeSpan.FromMilliseconds(1000)).Value;
        var addresses = new CellsRange(cellsRangeRef).Addresses();
        for (var i = 0; i < addresses.Count; i++)
        {
            var category = new SheetCategory(chartPart, sheetNameRef, addresses[i], cachedValues[i]);
            categoryList.Add(category);
        }

        return categoryList;
    }
}
