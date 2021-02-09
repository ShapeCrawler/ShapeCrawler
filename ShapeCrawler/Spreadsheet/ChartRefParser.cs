using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Charts;
using SlideDotNet.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Spreadsheet
{
    /// <summary>
    /// Represents a parser of series point value.
    /// </summary>
    internal class ChartRefParser
    {
        private readonly ChartSc _chart;

        internal ChartRefParser(ChartSc chart)
        {
            _chart = chart;
        }

        #region Public Methods

        /// <summary>
        /// Gets series values from xlsx.
        /// </summary>
        /// <param name="numberReference"></param>
        /// <param name="chartPart"></param>
        /// <returns></returns>
        internal IReadOnlyList<double> GetNumbers(C.NumberReference numberReference, ChartPart chartPart)
        {
            if (numberReference.NumberingCache != null)
            {
                // From cache
                IEnumerable<C.NumericValue> cNumericValues = numberReference.NumberingCache.Descendants<C.NumericValue>();
                var pointValues = new List<double>(cNumericValues.Count());
                foreach (var numericValue in cNumericValues)
                {
                    var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    var roundNumber = Math.Round(number, 1);
                    pointValues.Add(roundNumber);
                }

                return pointValues;
            }

            // From Spreadsheet
            List<string> cellStrValues = GetCellStrValues(numberReference.Formula, chartPart.EmbeddedPackagePart);
            var cellNumberValues = new List<double>(cellStrValues.Count); // TODO: consider allocate on stack
            foreach (string cellValue in cellStrValues)
            {
                string cellValueStr = cellValue;
                cellValueStr = cellValueStr.Length == 0 ? "0" : cellValueStr;
                cellNumberValues.Add(double.Parse(cellValueStr, CultureInfo.InvariantCulture.NumberFormat));
            }

            return cellNumberValues;
        }

        internal string GetSingleString(C.StringReference strRef, ChartPart chartPart)
        {
            var fromCache = strRef.StringCache?.GetFirstChild<C.StringPoint>().Single().InnerText;
            if (fromCache != null)
            {
                return fromCache;
            }

            var formula = strRef.Formula;
            var cellStrValues = GetCellStrValues(formula, chartPart.EmbeddedPackagePart);

            return cellStrValues.Single();
        }

        #endregion Public Methods

        #region Private Methods

        private List<string> GetCellStrValues(C.Formula formula, EmbeddedPackagePart xlsxPackagePart) //EmbeddedPackagePart : OpenXmlPart
        {
            var exist = _chart.Shape.Slide.Presentation.PresentationData.SpreadsheetCache.TryGetValue(xlsxPackagePart, out var xlsxDoc);
            if (!exist)
            {
                xlsxDoc = SpreadsheetDocument.Open(xlsxPackagePart.GetStream(), false);
                _chart.Shape.Slide.Presentation.PresentationData.SpreadsheetCache.Add(xlsxPackagePart, xlsxDoc);
            }

            string filteredFormula = GetFilteredFormula(formula);
            string[] sheetNameAndCellsFormula = filteredFormula.Split('!'); //eg: Sheet1!A2:A5 -> ['Sheet1', 'A2:A5']
            WorkbookPart workbookPart = xlsxDoc.WorkbookPart;
            string sheetId = workbookPart.Workbook.Sheets.Elements<Sheet>().First(sheet => sheetNameAndCellsFormula[0] == sheet.Name).Id;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            IEnumerable<Cell> cells = worksheetPart.Worksheet.GetFirstChild<SheetData>().ChildElements.SelectMany(r => r.Elements<Cell>()); //TODO: use HashSet
            var addresses = new CellFormulaParser(sheetNameAndCellsFormula[1]).GetCellAddresses(); //eg: [1] = 'A2:A5'

            var strValues = new List<string>(addresses.Count);
            foreach (var address in addresses)
            {
                var sdkCellValueStr = cells.First(c => c.CellReference == address).InnerText;
                strValues.Add(sdkCellValueStr);
            }

            return strValues;
        }

        private static string GetFilteredFormula(C.Formula formula)
        {
#if NETSTANDARD2_1 || NETCOREAPP2_0 || NET5_0
            var filteredFormula = formula.Text
                .Replace("'", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace("$", string.Empty,
                    StringComparison.OrdinalIgnoreCase); //eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5            
#else
            var filteredFormula = formula.Text.Replace("'", string.Empty).Replace("$", string.Empty);
#endif
            return filteredFormula;
        }

        #endregion Private Methods
    }
}