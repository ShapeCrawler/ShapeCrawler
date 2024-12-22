using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ShapeCrawler.Excel;

internal readonly record struct ExcelCellsRange
{
    private readonly string range;
    private readonly LinkedList<string> tempList = new ();

    #region Constructors

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelCellsRange"/> class.
    /// </summary>
    /// <param name="range">Cells range (eg. "A2:A5").</param>
    internal ExcelCellsRange(string range)
    {
        this.range = range;
    }

    #endregion Constructors

    /// <summary>
    ///     Gets collection of the cell's addresses like ['B10','B11','B12'].
    /// </summary>
    /// <remarks>input="B10:B12", output=['B10','B11','B12'].</remarks>
    internal List<string> Addresses()
    {
        this.Letter();
        return [.. this.tempList];
    }

    #region Private Methods

    private void Letter(int startIndex = 0)
    {
        var letterCharacters = this.range[startIndex..].TakeWhile(char.IsLetter);
        var letterStr = string.Concat(letterCharacters);
        var nextStart = startIndex + letterCharacters.Count();

        this.Digit(letterStr, nextStart);
    }

    private void Digit(string letterPart, int startIndex)
    {
        int digitInt = this.Digit(startIndex);
        this.tempList.AddLast(letterPart + digitInt); // e.g. 'B'+'10' -> B10

        int endIndex = startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length;
        if (endIndex >= this.range.Length)
        {
            return;
        }

        var nextStart = endIndex + letterPart.Length + 1; // skip separator and letter lengths
        if (this.range[endIndex] == ':')
        {
            this.LetterAfterColon(letterPart, digitInt, nextStart);
        }

        if (this.range[endIndex] == ',')
        {
            this.Letter(nextStart);
        }
    }

    private void LetterAfterColon(string letterPart, int digitPart, int startIndex)
    {
        var digitInt = this.Digit(startIndex);
        for (var nextDigitInt = digitPart + 1; nextDigitInt <= digitInt; nextDigitInt++)
        {
            this.tempList.AddLast(letterPart + nextDigitInt);
        }

        var nextStart =
            startIndex + digitInt.ToString(CultureInfo.CurrentCulture).Length +
            1; // skip last digit and separator characters
        if (nextStart < this.range.Length)
        {
            this.Letter(nextStart);
        }
    }

    private int Digit(int startIndex)
    {
        var digitChars = this.range[startIndex..].TakeWhile(char.IsDigit);
        return int.Parse(string.Concat(digitChars), CultureInfo.CurrentCulture);
    }

    #endregion Private Methods
}