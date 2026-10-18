using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal sealed class FontReplacement(PresentationDocument document)
{
    internal void Replace(string oldName, string newName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(oldName);
        ArgumentException.ThrowIfNullOrWhiteSpace(newName);
        if (oldName == newName)
        {
            return;
        }

        var visited = new HashSet<OpenXmlPart>();
        foreach (var relationship in document.Parts)
        {
            ReplaceInPart(relationship.OpenXmlPart, oldName, newName, visited);
        }
    }

    private static void ReplaceInPart(OpenXmlPart part, string oldName, string newName, HashSet<OpenXmlPart> visited)
    {
        if (!visited.Add(part))
        {
            return;
        }

        var root = part.RootElement;
        if (root is not null)
        {
            foreach (var font in root.Descendants<A.TextFontType>())
            {
                if (string.Equals(font.Typeface?.Value, oldName, StringComparison.Ordinal))
                {
                    font.Typeface = newName;
                }
            }

            foreach (var font in root.Descendants<A.SupplementalFont>())
            {
                if (string.Equals(font.Typeface?.Value, oldName, StringComparison.Ordinal))
                {
                    font.Typeface = newName;
                }
            }
        }

        foreach (var relationship in part.Parts)
        {
            ReplaceInPart(relationship.OpenXmlPart, oldName, newName, visited);
        }
    }
}
