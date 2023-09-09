using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class ExcelBook
{
    private readonly EmbeddedPackagePart sdkEmbeddedPackagePart;
    private Stream? embeddedPackagePartStream;

    internal ExcelBook(EmbeddedPackagePart sdkEmbeddedPackagePart)
    {
        this.sdkEmbeddedPackagePart = sdkEmbeddedPackagePart;
        this.SpreadsheetDocument = new Lazy<SpreadsheetDocument>(this.GetSpreadsheetDocument);
    }

    internal WorkbookPart WorkbookPart => this.SpreadsheetDocument.Value.WorkbookPart!;

    internal byte[] BinaryData => this.GetByteArray();

    internal Lazy<SpreadsheetDocument> SpreadsheetDocument { get; }
    
    private SpreadsheetDocument GetSpreadsheetDocument()
    {
        this.embeddedPackagePartStream = this.sdkEmbeddedPackagePart.GetStream();
        var spreadsheetDocument = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(this.embeddedPackagePartStream, true);

        return spreadsheetDocument;
    }

    private byte[] GetByteArray()
    {
        var mStream = new MemoryStream();
        this.SpreadsheetDocument.Value.Clone(mStream);

        return mStream.ToArray();
    }
}