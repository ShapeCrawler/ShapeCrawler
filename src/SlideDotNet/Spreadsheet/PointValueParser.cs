using System;
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
    /// TODO: convert into interface
    public class PointValueParser
    {
        public static IList<double> FromFormula(C.Formula formula, EmbeddedPackagePart embeddedPackagePart)
        {
            //TODO: caching embeddedPackagePart
            Check.NotNull(formula, nameof(formula));
            Check.NotNull(embeddedPackagePart, nameof(embeddedPackagePart));

            var filteredFormula = formula.Text.Replace("'", string.Empty, StringComparison.Ordinal)
                                                    .Replace("$", string.Empty, StringComparison.Ordinal);
            var sheetNameAndCellsFormula = filteredFormula.Split('!');
            var stream = embeddedPackagePart.GetStream();
            var doc = SpreadsheetDocument.Open(stream, false);
            var wbPart = doc.WorkbookPart;
            string sheetId = wbPart.Workbook.Descendants<Sheet>().First(s => sheetNameAndCellsFormula[0].Equals(s.Name, StringComparison.Ordinal)).Id;
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheetId);
            var sdkCells = wsPart.Worksheet.Descendants<Cell>(); //TODO: use HashSet
            var addresses = new CellFormulaParser(sheetNameAndCellsFormula[1]).GetCellAddresses();
            var result = new List<double>(addresses.Count);
            foreach (var address in addresses)
            {
                var sdkCellValueStr = sdkCells.First(c => c.CellReference == address).InnerText.Replace(".", ",", StringComparison.Ordinal);
                sdkCellValueStr = sdkCellValueStr == string.Empty ? "0" : sdkCellValueStr;
                result.Add(double.Parse(sdkCellValueStr));
            }

            doc.Close();
            stream.Close();
            return result;
        }

        public static IList<double> FromCache(C.NumberingCache numberingCache)
        {
            Check.NotNull(numberingCache, nameof(numberingCache));

            var sdkNumericValues = numberingCache.Descendants<C.NumericValue>();
            var pointValues = new List<double>(sdkNumericValues.Count());
            foreach (var numericValue in sdkNumericValues)
            {
                var sdkValue = numericValue.InnerText.Replace(".", ",", StringComparison.Ordinal); // double type uses comma as decimal separator
                pointValues.Add(double.Parse(sdkValue));
            }

            return pointValues;
        }

        public static IList<double> FromNumRef(C.NumberReference numRef, EmbeddedPackagePart xlsxPackagePart)
        {
            var numberingCache = numRef.NumberingCache;
            if (numberingCache != null)
            {
                return FromCache(numberingCache).ToList(); //TODO: remove ToList()
            }

            return FromFormula(numRef.Formula, xlsxPackagePart).ToList(); //TODO: remove ToList()
        }
    }
}