using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ShapeCrawler.Charts;

internal sealed class CellsRangeParser
{
    private readonly string cellRange;
    private readonly LinkedList<string> tempList = new ();

    #region Constructors

    /// <summary>
    ///     Initializes a new instance of the <see cref="CellsRangeParser"/> class.
    /// </summary>
    /// <param name="cellRange">Cells range (i.e. "A2:A5").</param>
    internal CellsRangeParser(string cellRange)
    {
        this.cellRange = cellRange;
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

    #region Private Methods

    private void ParseLetter(int startIndex = 0)
    {
        var letterCharacters = this.cellRange.Substring(startIndex).TakeWhile(char.IsLetter);
        var letterStr = string.Concat(letterCharacters);
        var nextStart = startIndex + letterCharacters.Count();

        this.ParseDigit(letterStr, nextStart);
    }

    private void ParseDigit(string letterPart, int startIndex)
    {
        int digitInt = this.GetDigit(startIndex);
        this.tempList.AddLast(letterPart + digitInt); // e.g. 'B'+'10' -> B10

        int endIndex = startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length;
        if (endIndex >= this.cellRange.Length)
        {
            return;
        }

        var nextStart = endIndex + letterPart.Length + 1; // skip separator and letter lengths
        if (this.cellRange[endIndex] == ':')
        {
            this.ParseLetterAfterColon(letterPart, digitInt, nextStart);
        }

        if (this.cellRange[endIndex] == ',')
        {
            this.ParseLetter(nextStart);
        }
    }

    private void ParseLetterAfterColon(string letterPart, int digitPart, int startIndex)
    {
        var digitInt = this.GetDigit(startIndex);
        for (var nextDigitInt = digitPart + 1; nextDigitInt <= digitInt; nextDigitInt++)
        {
            this.tempList.AddLast(letterPart + nextDigitInt);
        }

        var nextStart =
            startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length +
            1; // skip last digit and separator characters
        if (nextStart < this.cellRange.Length)
        {
            this.ParseLetter(nextStart);
        }
    }

    private int GetDigit(int startIndex)
    {
        var digitChars = this.cellRange.Substring(startIndex).TakeWhile(char.IsDigit);
        return int.Parse(string.Concat(digitChars), CultureInfo.CurrentCulture);
    }

    #endregion Private Methods
}