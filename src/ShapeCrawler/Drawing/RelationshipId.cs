using System;

namespace ShapeCrawler.Drawing;

internal struct RelationshipId
{
    internal static string New() => $"rId-{Guid.NewGuid().ToString("N")[..5]}";
}