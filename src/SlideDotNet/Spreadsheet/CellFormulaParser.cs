using System.Collections.Generic;
using System.Linq;
// ReSharper disable All

namespace SlideDotNet.Spreadsheet
{
    /// <summary>
    /// Represents a cell formula parser.
    /// </summary>
    public class CellFormulaParser
    {
        #region Fields

        private readonly string _cellsFormula;
        private readonly LinkedList<string> _tempList = new LinkedList<string>();

        #endregion Fields

        #region Constructors

        public CellFormulaParser(string cellsFormula)
        {
            _cellsFormula = cellsFormula;
        }

        #endregion Constructors

        /// <summary>
        /// Gets collection of the cell's addresses like ['B10','B11','B12'].
        /// </summary>
        /// <remarks>input="B10:B12", output=['B10','B11','B12']</remarks>
        public IList<string> GetCellAddresses()
        {
            ParseLetter();
            return _tempList.ToArray();
        }

        #region Private Methods

        private void ParseLetter(int startIndex = 0)
        {
            var letterCharacters = _cellsFormula.Substring(startIndex).TakeWhile(char.IsLetter);
            var letterStr = string.Concat(letterCharacters);
            var nextStart = startIndex + letterCharacters.Count();

            ParseDigit(letterStr, nextStart);
        }

        private void ParseDigit(string letterPart, int startIndex)
        {
            var digitInt = GetDigit(startIndex);
            _tempList.AddLast(letterPart + digitInt); // e.g. 'B'+'10' -> B10

            var endIndex = startIndex + digitInt.ToString().Length;
            if (endIndex >= _cellsFormula.Length)
            {
                return;
            }

            var nextStart = endIndex + letterPart.Length + 1; // skip separator and letter lengths
            if (_cellsFormula[endIndex] == ':')
            {
                ParseLetterAfterColon(letterPart, digitInt, nextStart);
            }
            if (_cellsFormula[endIndex] == ',')
            {
                ParseLetter(nextStart);
            }
        }

        private void ParseLetterAfterColon(string letterPart, int digitPart, int startIndex)
        {
            var digitInt = GetDigit(startIndex);
            for (var nextDigitInt = digitPart + 1; nextDigitInt <= digitInt; nextDigitInt++)
            {
                _tempList.AddLast(letterPart + nextDigitInt);
            }

            var nextStart = startIndex + digitInt.ToString().Length + 1; // skip last digit and separator characters
            if (nextStart < _cellsFormula.Length)
            {
                ParseLetter(nextStart);
            }
        }

        private int GetDigit(int startIndex)
        {
            var digitChars = _cellsFormula.Substring(startIndex).TakeWhile(char.IsDigit);
            return int.Parse(string.Concat(digitChars));
        }

        #endregion Private Methods
    }
}