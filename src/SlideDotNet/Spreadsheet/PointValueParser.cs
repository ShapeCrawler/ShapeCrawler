using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SlideDotNet.Validation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Spreadsheet
{
    /// <summary>
    /// Represents a series value point parser.
    /// </summary>
    public class PointValueParser
    {
        public static IList<double> FromFormula(C.Formula formula, EmbeddedPackagePart embeddedPackagePart)
        {
            //TODO: caching embeddedPackagePart
            Check.NotNull(formula, nameof(formula));
            Check.NotNull(embeddedPackagePart, nameof(embeddedPackagePart));

            var filteredFormula = formula.Text.Replace("'", string.Empty)
                                                    .Replace("$", string.Empty);
            var sheetNameAndCellsFormula = filteredFormula.Split('!');
            var stream = embeddedPackagePart.GetStream();
            var doc = SpreadsheetDocument.Open(stream, false);
            var wbPart = doc.WorkbookPart;
            string sheetId = wbPart.Workbook.Descendants<Sheet>().First(s => sheetNameAndCellsFormula[0].Equals(s.Name)).Id;
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
            var sdkCells = wsPart.Worksheet.Descendants<Cell>(); //TODO: use HashSet
            var addresses = new CellFormulaParser(sheetNameAndCellsFormula[1]).GetCellAddresses();
            var result = new List<double>(addresses.Count);
            foreach (var address in addresses)
            {
                var sdkCellValueStr = sdkCells.First(c => c.CellReference == address).InnerText.Replace(".", ",");
                sdkCellValueStr = sdkCellValueStr == string.Empty ? "0" : sdkCellValueStr;
                result.Add(double.Parse(sdkCellValueStr));
            }

            doc.Close();
            stream.Close();
            return result;
        }

        public static IList<double> FromCache(C.NumberingCache numberingCache)
        {
            var sdkNumericValues = numberingCache.Descendants<C.NumericValue>();
            var pointValues = new List<double>(sdkNumericValues.Count());
            foreach (var numericValue in sdkNumericValues)
            {
                var sdkValue = numericValue.InnerText.Replace(".", ","); // double type uses comma as decimal separator
                pointValues.Add(double.Parse(sdkValue));
            }

            return pointValues;
        }
    }
}