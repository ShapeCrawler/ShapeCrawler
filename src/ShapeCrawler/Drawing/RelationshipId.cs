using System;

namespace ShapeCrawler.Drawing;

internal struct RelationshipId
{
    public string New()
    {
#if NETSTANDARD2_0
        return $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
#else
        return $"rId-{Guid.NewGuid().ToString("N")[..5]}";
#endif
    }
}