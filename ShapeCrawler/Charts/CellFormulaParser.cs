using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ShapeCrawler.Spreadsheet
{
    /// <summary>
    ///     Represents the cell formula parser.
    /// </summary>
    internal class CellFormulaParser
    {
        #region Constructors

        internal CellFormulaParser(string formula)
        {
            this.formula = formula;
        }

        #endregion Constructors

        /// <summary>
        ///     Gets collection of the cell's addresses like ['B10','B11','B12'].
        /// </summary>
        /// <remarks>input="B10:B12", output=['B10','B11','B12'].</remarks>
        internal List<string> GetCellAddresses()
        {
            this.ParseLetter();

            return this.tempList.ToList();
        }

        #region Fields

        private readonly string formula;
        private readonly LinkedList<string> tempList = new LinkedList<string>();

        #endregion Fields

        #region Private Methods

        private void ParseLetter(int startIndex = 0)
        {
            var letterCharacters = this.formula.Substring(startIndex).TakeWhile(char.IsLetter);
            var letterStr = string.Concat(letterCharacters);
            var nextStart = startIndex + letterCharacters.Count();

            ParseDigit(letterStr, nextStart);
        }

        private void ParseDigit(string letterPart, int startIndex)
        {
            int digitInt = this.GetDigit(startIndex);
            tempList.AddLast(letterPart + digitInt); // e.g. 'B'+'10' -> B10

            int endIndex = startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length;
            if (endIndex >= formula.Length)
            {
                return;
            }

            var nextStart = endIndex + letterPart.Length + 1; // skip separator and letter lengths
            if (formula[endIndex] == ':')
            {
                ParseLetterAfterColon(letterPart, digitInt, nextStart);
            }

            if (formula[endIndex] == ',')
            {
                ParseLetter(nextStart);
            }
        }

        private void ParseLetterAfterColon(string letterPart, int digitPart, int startIndex)
        {
            var digitInt = GetDigit(startIndex);
            for (var nextDigitInt = digitPart + 1; nextDigitInt <= digitInt; nextDigitInt++)
            {
                tempList.AddLast(letterPart + nextDigitInt);
            }

            var nextStart =
                startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length +
                1; // skip last digit and separator characters
            if (nextStart < formula.Length)
            {
                ParseLetter(nextStart);
            }
        }

        private int GetDigit(int startIndex)
        {
            IEnumerable<char> digitChars = formula.Substring(startIndex).TakeWhile(char.IsDigit);
            return int.Parse(string.Concat(digitChars), CultureInfo.CurrentCulture);
        }

        #endregion Private Methods
    }
}