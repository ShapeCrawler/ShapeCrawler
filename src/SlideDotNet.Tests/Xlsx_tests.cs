using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using CellFormula = SlideDotNet.Spreadsheet.CellFormula;

// ReSharper disable PossibleMultipleEnumeration

namespace SlideDotNet.Tests
{


    public class Xlsx_tests
    {
        [Fact]
        public void Workbook_Test()
        {
            var formula1 = @"'ShareMonitor figures'!$B$11:$B$846";
            var formula2 = @"Sheet1!$B$11:$B$846";
            var formula3 = @"(Sheet1!$A$1:$A$3,Sheet1!$A$5)";

            string pattern = @"(?<=')(.*?)(?=')";
            Regex rg = new Regex(pattern);
            MatchCollection matchedCollection = rg.Matches(formula1);
            var sheetName = matchedCollection.Single().Value;
            var filePath = @"d:\test.xlsx";
            var spreadDoc = SpreadsheetDocument.Open(filePath, false);

            WorkbookPart workbookPart = spreadDoc.WorkbookPart;
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            var sheetPart = (WorksheetPart)workbookPart.GetPartById(relId);
            var sheetData = sheetPart.Worksheet.Elements<SheetData>().First();
            
            spreadDoc.Close();
        }

        [Fact]
        public void Parse_Test()
        {
            // Arrange
            var testExpression = "B10:B12";
            var cellRangeParser = new CellFormula(testExpression);
            var testExpression2 = "B10:B12,B14";
            var cellRangeParser2 = new CellFormula(testExpression2);
            var testExpression3 = "B10";
            var cellRangeParser3 = new CellFormula(testExpression3);

            // Act
            var result = cellRangeParser.GetCellAddresses();
            var result2 = cellRangeParser2.GetCellAddresses();
            var result3 = cellRangeParser3.GetCellAddresses();

            // Assert
            Assert.Equal(@"B10", result[0]);
            Assert.Equal(@"B11", result[1]);
            Assert.Equal(@"B12", result[2]);
            Assert.Equal(@"B14", result2[3]);
            Assert.Equal(@"B10", result3.Single());
        }
    }
}
