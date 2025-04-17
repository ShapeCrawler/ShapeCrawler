using System;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

internal class DefaultPackageProperties : IPackageProperties
{
    public string? Category { get; set; }

    public string? ContentStatus { get; set; }

    public DateTime? Created { get; set; } = DateTime.Now;

    public string? Creator { get; set; }

    public string? Description { get; set; }

    public string? Identifier { get; set; }

    public string? Keywords { get; set; }

    public string? Language { get; set; }

    public string? LastModifiedBy { get; set; }

    public DateTime? LastPrinted { get; set; }

    public DateTime? Modified { get; set; } = DateTime.Now;

    public string? Revision { get; set; } = "1";

    public string? Subject { get; set; }

    public string? Title { get; set; }

    public string? Version { get; set; } = "1.0";

    public string? ContentType { get; set; } = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
}
