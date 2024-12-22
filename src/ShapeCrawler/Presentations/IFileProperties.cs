using System;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the core properties of the presentation file.
/// </summary>
/// <remarks>
///     These properties are not related to the presentation itself, but rather the containing file.
/// </remarks>
public interface IFileProperties
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
    ///     Gets or sets the primary creator. The identification is environment-specific and can consist of a name, email address, employee ID, etc. It is recommended that this value be only as verbose as necessary to identify the individual.
    /// </summary>
    string? Author { get; set; }

    /// <summary>
    ///     Gets or sets the description or abstract of the contents.
    /// </summary>
    string? Comments { get; set; }

    /// <summary>
    ///     Gets or sets a delimited set of keywords (tags) to support searching and indexing. This is typically a list of terms that are not available elsewhere in the properties.
    /// </summary>
    /// <remarks>
    ///     The delimeter to use is not specified.
    /// </remarks>
    string? Tags { get; set; }

    /// <summary>
    ///     Gets or sets the primary language of the package content. The language tag is composed of one or more parts: A primary language subtag and a (possibly empty) series of subsequent subtags, for example, "EN-US". These values MUST follow the convention specified in RFC 3066.
    /// </summary>
    /// <remarks>
    ///     Show in File Explorer, but not in PowerPoint client.
    /// </remarks>
    string? Language { get; set; }

    /// <summary>
    ///     Gets or sets the user who performed the last modification. The identification is environment-specific and can consist of a name, email address, employee ID, etc. It is recommended that this value be only as verbose as necessary to identify the individual.
    /// </summary>
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
    ///     Gets or sets the revision number. This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision.
    /// </summary>
    /// <remarks>
    ///     Show in File Explorer, but not in PowerPoint client.
    /// </remarks>
    int? RevisionNumber { get; set; }

    /// <summary>
    ///     Gets or sets the topic of the contents.
    /// </summary>
    string? Subject { get; set; }

    /// <summary>
    ///     Gets or sets the title.
    /// </summary>
    string? Title { get; set; }

    /// <summary>
    ///     Gets or sets the version number. This value is set by the user or by the application.
    /// </summary>
    string? Version { get; set; }
}