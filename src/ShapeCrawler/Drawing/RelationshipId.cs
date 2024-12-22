using System;

namespace ShapeCrawler.Drawing;

internal struct RelationshipId
{
    public string New()
    {
        return $"rId-{Guid.NewGuid().ToString("N")[..5]}";
    }
}