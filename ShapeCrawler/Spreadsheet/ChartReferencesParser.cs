using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Spreadsheet
{
    internal class ChartReferencesParser
    {
        #region Private Methods

        private static string GetFilteredFormula(C.Formula formula)
        {
#if NETSTANDARD2_1 || NET5_0 || NETCOREAPP2_1
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

        #region Internal Methods

        internal static IReadOnlyList<double> GetNumbersFromCacheOrSpreadsheet(C.NumberReference numberReference,
            SCChart slideChart)
        {
            if (numberReference.NumberingCache != null)
            {
                // From cache
                IEnumerable<C.NumericValue> cNumericValues =
                    numberReference.NumberingCache.Descendants<C.NumericValue>();
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
            List<X.Cell> xCells = GetXCellsByFormula(numberReference.Formula, slideChart);
            var cellNumberValues = new List<double>(xCells.Count); // TODO: consider allocate on stack
            foreach (X.Cell xCell in xCells)
            {
                string cellValueStr = xCell.InnerText;
                cellValueStr = cellValueStr.Length == 0 ? "0" : cellValueStr;
                cellNumberValues.Add(double.Parse(cellValueStr, CultureInfo.InvariantCulture.NumberFormat));
            }

            return cellNumberValues;
        }

        internal static string GetSingleString(C.StringReference stringReference, SCChart slideChart)
        {
            string fromCache = stringReference.StringCache?.GetFirstChild<C.StringPoint>().Single().InnerText;
            if (fromCache != null)
            {
                return fromCache;
            }

            List<X.Cell> xCell = GetXCellsByFormula(stringReference.Formula, slideChart);

            return xCell.Single().InnerText;
        }

        /// <summary>
        ///     Gets cell values.
        /// </summary>
        /// <param name="cFormula">
        ///     Cell range formula (c:f).
        ///     <c:cat>
        ///         <c:strRef>
        ///             <c:f>
        ///                 Sheet1!$A$2:$A$3
        ///             </c:f>
        ///         </c:strRef>
        ///     </c:cat>
        /// </param>
        /// <param name="slideChart"></param>
        internal static List<X.Cell> GetXCellsByFormula(C.Formula cFormula, SCChart slideChart)
        {
            // Get all <x:c> elements of formula sheet
            string filteredFormula = GetFilteredFormula(cFormula);
            string[] sheetNameAndCellsRange = filteredFormula.Split('!'); //eg: Sheet1!A2:A5 -> ['Sheet1', 'A2:A5']
            WorkbookPart workbookPart = slideChart.ChartWorkbook.WorkbookPart;
            string chartSheetName = sheetNameAndCellsRange[0];
            string chartSheetId = workbookPart.Workbook.Sheets.Elements<X.Sheet>()
                .First(xSheet => xSheet.Name == chartSheetName).Id;
            WorksheetPart worksheetPart = (WorksheetPart) workbookPart.GetPartById(chartSheetId);
            IEnumerable<X.Cell> allXCells = worksheetPart.Worksheet.GetFirstChild<X.SheetData>().ChildElements
                .SelectMany(e => e.Elements<X.Cell>()); //TODO: use HashSet

            List<string> formulaCellAddressList = new CellFormulaParser(sheetNameAndCellsRange[1]).GetCellAddresses();

            List<X.Cell> xCells = new (formulaCellAddressList.Count);
            foreach (string address in formulaCellAddressList)
            {
                X.Cell xCell = allXCells.First(cell => cell.CellReference == address);
                xCells.Add(xCell);
            }

            return xCells;
        }

        #endregion Internal Methods
    }
}