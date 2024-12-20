using System;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Presentations;

internal class FileProperties: IFileProperties
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