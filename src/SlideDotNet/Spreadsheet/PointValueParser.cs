using System;
using System.Collections.Generic;
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
        public static IList<double> FromNumRef(C.NumberReference numRef, EmbeddedPackagePart xlsxPackagePart)
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
                    var sdkValue = numericValue.InnerText.Replace(".", ",", StringComparison.Ordinal); // double type uses comma as decimal separator
                    pointValues.Add(double.Parse(sdkValue));
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

            throw new NotImplementedException();
        }

        #endregion Public Methods

        #region Private Methods

        private static IList<double> FromFormula(C.Formula formula, EmbeddedPackagePart xlsxPackagePart)
        {
            //TODO: caching embeddedPackagePart
            var filteredFormula = formula.Text.Replace("'", string.Empty, StringComparison.Ordinal)
                                                    .Replace("$", string.Empty, StringComparison.Ordinal); //eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
            var sheetNameAndCellsFormula = filteredFormula.Split('!'); //eg: Sheet1!A2:A5 -> ['Sheet1', 'A2:A5']
            var stream = xlsxPackagePart.GetStream();
            var doc = SpreadsheetDocument.Open(stream, false);
            var wbPart = doc.WorkbookPart;
            string sheetId = wbPart.Workbook.Descendants<Sheet>().First(s => sheetNameAndCellsFormula[0].Equals(s.Name, StringComparison.Ordinal)).Id;
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
            var sdkCells = wsPart.Worksheet.Descendants<Cell>(); //TODO: use HashSet
            var addresses = new CellFormulaParser(sheetNameAndCellsFormula[1]).GetCellAddresses(); //eg: [1] = 'A2:A5'
            
            var result = new List<double>(addresses.Count);
            foreach (var address in addresses)
            {
                var sdkCellValueStr = sdkCells.First(c => c.CellReference == address).InnerText
                                                                            .Replace(".", ",", StringComparison.Ordinal);
                sdkCellValueStr = sdkCellValueStr.Length == 0 ? "0" : sdkCellValueStr;
                result.Add(double.Parse(sdkCellValueStr));
            }

            doc.Close();
            stream.Close();
            return result;
        }

        #endregion Private Methods
    }
}