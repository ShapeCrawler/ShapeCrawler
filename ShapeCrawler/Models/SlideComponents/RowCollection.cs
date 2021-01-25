using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Collections;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
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
        private readonly Dictionary<RowSc, A.TableRow> _innerSdkDic;

        #region Constructors

        public RowCollection(IEnumerable<A.TableRow> sdkTblRows)
        {
            Check.NotNull(sdkTblRows, nameof(sdkTblRows));

            var count = sdkTblRows.Count();
            CollectionItems = new List<RowSc>(count);
            _innerSdkDic = new Dictionary<RowSc, A.TableRow>(count);
            foreach (var sdkRow in sdkTblRows)
            {
                var innerRow = new RowSc(sdkRow);

                _innerSdkDic.Add(innerRow, sdkRow);
                CollectionItems.Add(innerRow);
            }
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Removes the specified table row.
        /// </summary>
        /// <param name="item"></param>
        public override void Remove(RowSc item)
        {
            if (!_innerSdkDic.ContainsKey(item))
            {
                throw new ArgumentNullException(nameof(item));
            }

            _innerSdkDic[item].Remove();
            CollectionItems.Remove(item);
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