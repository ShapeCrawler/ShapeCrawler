using System;
using DocumentFormat.OpenXml.Packaging;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the metadata of the presentation file.
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

internal class FileProperties : IPresentationProperties
{
    private readonly DocumentFormat.OpenXml.Packaging.IPackageProperties sdkPackageProperties;

    internal FileProperties(CoreFilePropertiesPart sdkPart)
    {
        this.sdkPackageProperties = sdkPart.OpenXmlPackage.PackageProperties;
    }

    public string? Author
    {
        get => this.sdkPackageProperties.Creator;
        set => this.sdkPackageProperties.Creator = value;
    }

    public string? Categories 
    {
        get => this.sdkPackageProperties.Category;
        set => this.sdkPackageProperties.Category = value;
    }
    
    public string? ContentType 
    {
        get => this.sdkPackageProperties.ContentType;
        set => this.sdkPackageProperties.ContentType = value;
    }
    
    public string? ContentStatus 
    {
        get => this.sdkPackageProperties.ContentStatus;
        set => this.sdkPackageProperties.ContentStatus = value;
    }
    
    public DateTime? Created 
    {
        get => this.sdkPackageProperties.Created;
        set => this.sdkPackageProperties.Created = value;
    }
    
    public string? Comments 
    {
        get => this.sdkPackageProperties.Description;
        set => this.sdkPackageProperties.Description = value;
    }
    
    public string? Identifier 
    {
        get => this.sdkPackageProperties.Identifier;
        set => this.sdkPackageProperties.Identifier = value;
    }
    
    public string? Tags 
    {
        get => this.sdkPackageProperties.Keywords;
        set => this.sdkPackageProperties.Keywords = value;
    }
    
    public string? Language 
    {
        get => this.sdkPackageProperties.Language;
        set => this.sdkPackageProperties.Language = value;
    }
    
    public string? LastModifiedBy 
    {
        get => this.sdkPackageProperties.LastModifiedBy;
        set => this.sdkPackageProperties.LastModifiedBy = value;
    }
    
    public DateTime? LastPrinted 
    {
        get => this.sdkPackageProperties.LastPrinted;
        set => this.sdkPackageProperties.LastPrinted = value;
    }
    
    public DateTime? Modified 
    {
        get => this.sdkPackageProperties.Modified;
        set => this.sdkPackageProperties.Modified = value;
    }
    
    public int? RevisionNumber 
    {
        get
        {
            var revision = this.sdkPackageProperties.Revision;
            if (string.IsNullOrWhiteSpace(revision) || !int.TryParse(revision, out var result))
            {
                return null;
            }
            else
            {
                return result;
            }
        }
        set => this.sdkPackageProperties.Revision = value?.ToString();
    }
    
    public string? Subject 
    {
        get => this.sdkPackageProperties.Subject;
        set => this.sdkPackageProperties.Subject = value;
    }
    
    public string? Title 
    {
        get => this.sdkPackageProperties.Title;
        set => this.sdkPackageProperties.Title = value;
    }
    
    public string? Version 
    {
        get => this.sdkPackageProperties.Version;
        set => this.sdkPackageProperties.Version = value;
    }    
}