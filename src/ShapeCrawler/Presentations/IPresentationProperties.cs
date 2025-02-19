using System;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the presentation properties.
/// </summary>
public interface IPresentationProperties
{
    /// <summary>
    ///     Gets or sets the categories.
    /// </summary>
    /// <remarks>
    ///     The method to delimit categories is not specified.
    /// </remarks>
    string? Categories { get; set; }

    /// <summary>
    ///     Gets or sets the status of the content. Example values include "Draft", "Reviewed", and "Final".
    /// </summary>
    string? ContentStatus { get; set; }

    /// <summary>
    ///     Gets or sets the creation date and time.
    /// </summary>
    DateTime? Created { get; set; }

    /// <summary>
    ///     Gets or sets the primary creator.
    /// </summary>
    /// <remarks>
    ///     The identification is environment-specific and can consist of a name, email address, employee ID, etc. It is recommended that this value be only as verbose as necessary to identify the individual.     
    /// </remarks>
    string? Author { get; set; }

    /// <summary>
    ///     Gets or sets the description or abstract of the contents.
    /// </summary>
    string? Comments { get; set; }

    /// <summary>
    ///     Gets or sets a delimited set of keywords (tags) to support searching and indexing.
    /// </summary>
    /// <remarks>
    ///      This is typically a list of terms that are not available elsewhere in the properties. The delimiter to use is not specified.
    /// </remarks>
    string? Tags { get; set; }

    /// <summary>
    ///     Gets or sets the primary language of the package content.
    /// </summary>
    /// <remarks>
    ///     The language tag is composed of one or more parts: A primary language subtag and a (possibly empty) series of subsequent subtags, for example, "EN-US". These values MUST follow the convention specified in RFC 3066. Show in File Explorer, but not in PowerPoint client.
    /// </remarks>
    string? Language { get; set; }

    /// <summary>
    ///     Gets or sets the user who performed the last modification.
    /// </summary>
    /// <remarks>
    ///     The identification is environment-specific and can consist of a name, email address, employee ID, etc. It is recommended that this value be only as verbose as necessary to identify the individual.
    /// </remarks>
    string? LastModifiedBy { get; set; }

    /// <summary>
    ///     Gets or sets the date and time of the last printing.
    /// </summary>
    DateTime? LastPrinted { get; set; }

    /// <summary>
    ///     Gets or sets the date and time of the last modification.
    /// </summary>
    DateTime? Modified { get; set; }

    /// <summary>
    ///     Gets or sets the revision number.
    /// </summary>
    /// <remarks>
    ///      This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision. Show in File Explorer, but not in PowerPoint client.
    /// </remarks>
    int? RevisionNumber { get; set; }

    /// <summary>
    ///     Gets or sets the topic of the content.
    /// </summary>
    string? Subject { get; set; }

    /// <summary>
    ///     Gets or sets the title.
    /// </summary>
    string? Title { get; set; }

    /// <summary>
    ///     Gets or sets the version number.
    /// </summary>
    /// <remarks>
    ///     This value is set by the user or by the application.
    /// </remarks>
    string? Version { get; set; }
}

internal class PresentationProperties(IPackageProperties packageProperties): IPresentationProperties
{
    public string? Author
    {
        get => packageProperties.Creator;
        set => packageProperties.Creator = value;
    }

    public string? Categories 
    {
        get => packageProperties.Category;
        set => packageProperties.Category = value;
    }
    
    public string? ContentType 
    {
        get => packageProperties.ContentType;
        set => packageProperties.ContentType = value;
    }
    
    public string? ContentStatus 
    {
        get => packageProperties.ContentStatus;
        set => packageProperties.ContentStatus = value;
    }
    
    public DateTime? Created 
    {
        get => packageProperties.Created;
        set => packageProperties.Created = value;
    }
    
    public string? Comments 
    {
        get => packageProperties.Description;
        set => packageProperties.Description = value;
    }
    
    public string? Identifier 
    {
        get => packageProperties.Identifier;
        set => packageProperties.Identifier = value;
    }
    
    public string? Tags 
    {
        get => packageProperties.Keywords;
        set => packageProperties.Keywords = value;
    }
    
    public string? Language 
    {
        get => packageProperties.Language;
        set => packageProperties.Language = value;
    }
    
    public string? LastModifiedBy 
    {
        get => packageProperties.LastModifiedBy;
        set => packageProperties.LastModifiedBy = value;
    }
    
    public DateTime? LastPrinted 
    {
        get => packageProperties.LastPrinted;
        set => packageProperties.LastPrinted = value;
    }
    
    public DateTime? Modified 
    {
        get => packageProperties.Modified;
        set => packageProperties.Modified = value;
    }
    
    public int? RevisionNumber 
    {
        get
        {
            var revision = packageProperties.Revision;
            if (string.IsNullOrWhiteSpace(revision) || !int.TryParse(revision, out var result))
            {
                return null;
            }
            else
            {
                return result;
            }
        }
        set => packageProperties.Revision = value?.ToString();
    }
    
    public string? Subject 
    {
        get => packageProperties.Subject;
        set => packageProperties.Subject = value;
    }
    
    public string? Title 
    {
        get => packageProperties.Title;
        set => packageProperties.Title = value;
    }
    
    public string? Version 
    {
        get => packageProperties.Version;
        set => packageProperties.Version = value;
    }    
}