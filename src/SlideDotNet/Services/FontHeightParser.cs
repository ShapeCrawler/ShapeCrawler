using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents font height parser.
    /// </summary>
    public static class FontHeightParser
    {
        /// <summary>
        /// Parses and returns font height from <see cref="OpenXmlCompositeElement"/> instance.
        /// </summary>
        /// <param name="compositeElement"></param>
        /// <returns></returns>
        public static Dictionary<int, int> FromCompositeElement(OpenXmlCompositeElement compositeElement)
        {
            var result = new Dictionary<int, int>();
            foreach (var textPr in compositeElement.Elements().Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal))) // <a:lvl1pPr>, <a:lvl2pPr>, etc.
            {
                var fs = textPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>()?.FontSize;
                if (fs == null)
                {
                    continue;
                }
                // fourth character of LocalName contains level number, example: "lvl1pPr -> 1, lvl2pPr -> 2, etc."
                var lvl = int.Parse(textPr.LocalName[3].ToString());
                result.Add(lvl, fs.Value);
            }

            return result;
        }
    }
}