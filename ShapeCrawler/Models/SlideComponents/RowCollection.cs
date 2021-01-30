using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Collections;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler.Models.SlideComponents
{
    /// <summary>
    /// Represents a table rows collection.
    /// </summary>
    public class RowCollection : EditableCollection<RowSc>
    {
        private readonly Dictionary<RowSc, A.TableRow> _rowToATblRow;

        #region Constructors

        public RowCollection(IEnumerable<A.TableRow> aTableRows)
        {
            var count = aTableRows.Count();
            CollectionItems = new List<RowSc>(count);
            _rowToATblRow = new Dictionary<RowSc, A.TableRow>(count);
            foreach (A.TableRow aTblRow in aTableRows)
            {
                var row = new RowSc(aTblRow);

                _rowToATblRow.Add(row, aTblRow);
                CollectionItems.Add(row);
            }
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Removes the specified table row.
        /// </summary>
        /// <param name="row"></param>
        public override void Remove(RowSc row)
        {
            _rowToATblRow[row].Remove();
            CollectionItems.Remove(row);
        }

        /// <summary>
        /// Removes table row by index.
        /// </summary>
        /// <param name="index"></param>
        public void RemoveAt(int index)
        {
            if (index < 0 || index >= CollectionItems.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            var innerRow = CollectionItems[index];
            Remove(innerRow);
        }

        #endregion Public Methods
    }
}