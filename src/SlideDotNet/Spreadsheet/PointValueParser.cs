using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SlideDotNet.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Spreadsheet
{
    /// <summary>
    /// Represents a series value point parser.
    /// </summary>
    /// TODO: convert into interface
    public class PointValueParser
    {
        #region Public Methods

        /// <summary>
        /// Gets series values from xls.
        /// </summary>
        /// <param name="numRef"></param>
        /// <param name="xlsxPackagePart"></param>
        /// <returns></returns>
        public static IList<double> GetNumbers(C.NumberReference numRef, EmbeddedPackagePart xlsxPackagePart)
        {
            Check.NotNull(numRef, nameof(numRef));
            Check.NotNull(xlsxPackagePart, nameof(xlsxPackagePart));

            var numberingCache = numRef.NumberingCache;
            if (numberingCache != null)
            {
                var sdkNumericValues = numberingCache.Descendants<C.NumericValue>();
                var pointValues = new List<double>(sdkNumericValues.Count());
                foreach (var numericValue in sdkNumericValues)
                {
                    var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    var roundNumber = Math.Round(number, 1);
                    pointValues.Add(roundNumber);
                }

                return pointValues;
            }

            return FromFormula(numRef.Formula, xlsxPackagePart).ToList(); //TODO: remove ToList()
        }

        public static string GetSingleString(C.StringReference strRef, EmbeddedPackagePart xlsxPackagePart)
        {
            var fromCache = strRef.StringCache?.GetFirstChild<C.StringPoint>().Single().InnerText;
            if (fromCache != null)
            {
                return fromCache;
            }

            var formula = strRef.Formula;
            var cellStrValues = GetCellStrValues(formula, xlsxPackagePart);

            return cellStrValues.Single();
        }

        #endregion Public Methods

        #region Private Methods

        private static List<double> FromFormula(C.Formula formula, OpenXmlPart xlsxPackagePart)
        {
            var cellStrValues = GetCellStrValues(formula, xlsxPackagePart);

            var cellNumberValues = new List<double>(cellStrValues.Count);
            foreach (var cellValue in cellStrValues)
            {
                var sdkCellValueStr = cellValue;
                sdkCellValueStr = sdkCellValueStr.Length == 0 ? "0" : sdkCellValueStr;
                cellNumberValues.Add(double.Parse(sdkCellValueStr, CultureInfo.InvariantCulture.NumberFormat));
            }

            return cellNumberValues;
        }

        private static List<string> GetCellStrValues(C.Formula formula, OpenXmlPart xlsxPackagePart) //EmbeddedPackagePart : OpenXmlPart
        {
            //TODO: caching embeddedPackagePart
            var filteredFormula = formula.Text.Replace("'", string.Empty, StringComparison.Ordinal)
                .Replace("$", string.Empty, StringComparison.Ordinal); //eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
            var sheetNameAndCellsFormula = filteredFormula.Split('!'); //eg: Sheet1!A2:A5 -> ['Sheet1', 'A2:A5']
            var xlsxDoc = SpreadsheetDocument.Open(xlsxPackagePart.GetStream(), false);
            var wbPart = xlsxDoc.WorkbookPart;
            string sheetId = wbPart.Workbook.Descendants<Sheet>().First(s => sheetNameAndCellsFormula[0].Equals(s.Name, StringComparison.Ordinal)).Id;
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
            var sdkCells = wsPart.Worksheet.Descendants<Cell>(); //TODO: use HashSet
            var addresses = new CellFormulaParser(sheetNameAndCellsFormula[1]).GetCellAddresses(); //eg: [1] = 'A2:A5'

            var strValues = new List<string>(addresses.Count);
            foreach (var address in addresses)
            {
                var sdkCellValueStr = sdkCells.First(c => c.CellReference == address).InnerText;
                strValues.Add(sdkCellValueStr);
            }

            xlsxDoc.Close();
            return strValues;
        }

        #endregion Private Methods
    }
}