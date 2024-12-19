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

    public string? Creator
    {
        get => sdkPackageProperties.Creator;
        set => sdkPackageProperties.Creator = value;
    }

    public string? Category 
    {
        get => sdkPackageProperties.Category;
        set => sdkPackageProperties.Category = value;
    }
    
    public string? ContentType 
    {
        get => sdkPackageProperties.ContentType;
        set => sdkPackageProperties.ContentType = value;
    }
    
    public string? ContentStatus 
    {
        get => sdkPackageProperties.ContentStatus;
        set => sdkPackageProperties.ContentStatus = value;
    }
    
    public DateTime? Created 
    {
        get => sdkPackageProperties.Created;
        set => sdkPackageProperties.Created = value;
    }
    
    public string? Description 
    {
        get => sdkPackageProperties.Description;
        set => sdkPackageProperties.Description = value;
    }
    
    public string? Identifier 
    {
        get => sdkPackageProperties.Identifier;
        set => sdkPackageProperties.Identifier = value;
    }
    
    public string? Keywords 
    {
        get => sdkPackageProperties.Keywords;
        set => sdkPackageProperties.Keywords = value;
    }
    
    public string? Language 
    {
        get => sdkPackageProperties.Language;
        set => sdkPackageProperties.Language = value;
    }
    
    public string? LastModifiedBy 
    {
        get => sdkPackageProperties.LastModifiedBy;
        set => sdkPackageProperties.LastModifiedBy = value;
    }
    
    public DateTime? LastPrinted 
    {
        get => sdkPackageProperties.LastPrinted;
        set => sdkPackageProperties.LastPrinted = value;
    }
    
    public DateTime? Modified 
    {
        get => sdkPackageProperties.Modified;
        set => sdkPackageProperties.Modified = value;
    }
    
    public string? Revision 
    {
        get => sdkPackageProperties.Revision;
        set => sdkPackageProperties.Revision = value;
    }
    
    public string? Subject 
    {
        get => sdkPackageProperties.Subject;
        set => sdkPackageProperties.Subject = value;
    }
    
    public string? Title 
    {
        get => sdkPackageProperties.Title;
        set => sdkPackageProperties.Title = value;
    }
    
    public string? Version 
    {
        get => sdkPackageProperties.Version;
        set => sdkPackageProperties.Version = value;
    }    
}