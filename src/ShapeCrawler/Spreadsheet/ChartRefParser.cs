using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using SlideDotNet.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Spreadsheet
{
    /// <summary>
    /// Represents a series value point parser.
    /// </summary>
    /// TODO: convert into interface
    public class ChartRefParser: IChartRefParser
    {
        private readonly IShapeContext _spContext;

        public ChartRefParser(IShapeContext spContext)
        {
            _spContext = spContext;
        }

        #region Public Methods

        /// <summary>
        /// Gets series values from xlsx.
        /// </summary>
        /// <param name="numRef"></param>
        /// <param name="chartPart"></param>
        /// <returns></returns>
        public IList<double> GetNumbers(C.NumberReference numRef, ChartPart chartPart)
        {
            Check.NotNull(numRef, nameof(numRef));
            Check.NotNull(chartPart, nameof(chartPart));

            var numberingCache = numRef.NumberingCache;
            if (numberingCache != null) // From cache
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

            return FromFormula(numRef.Formula, chartPart.EmbeddedPackagePart).ToList(); //TODO: remove ToList()
        }

        public string GetSingleString(C.StringReference strRef, ChartPart chartPart)
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

        private List<double> FromFormula(C.Formula formula, OpenXmlPart xlsxPackagePart)
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

        private List<string> GetCellStrValues(C.Formula formula, OpenXmlPart xlsxPackagePart) //EmbeddedPackagePart : OpenXmlPart
        {
            var exist = _spContext.presentationData.XlsxDocuments.TryGetValue(xlsxPackagePart, out var xlsxDoc);
            if (!exist)
            {
                xlsxDoc = SpreadsheetDocument.Open(xlsxPackagePart.GetStream(), false);
                _spContext.presentationData.XlsxDocuments.Add(xlsxPackagePart, xlsxDoc);
            }
#if NETSTANDARD2_1 || NETCOREAPP2_0
            var filteredFormula = formula.Text
                .Replace("'", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace("$", string.Empty, StringComparison.OrdinalIgnoreCase); //eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5            
#else
            var filteredFormula = formula.Text.Replace("'", string.Empty).Replace("$", string.Empty);
#endif
            

            var sheetNameAndCellsFormula = filteredFormula.Split('!'); //eg: Sheet1!A2:A5 -> ['Sheet1', 'A2:A5']
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