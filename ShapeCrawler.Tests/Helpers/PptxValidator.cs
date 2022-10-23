using System.Collections.Generic;
using SharpCompress.Readers;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using SharpCompress.Common;

namespace ShapeCrawler.Tests.Helpers;

public static class PptxValidator
{
    public static ValidateResponse Validate(MemoryStream pptxStream)
    {
        pptxStream.Position = 0;
        using IReader reader = ReaderFactory.Open(pptxStream);
        while (reader.MoveToNextEntry())
        {
            string NotesMasterType = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml";
                
            // Validate Content_Type.xml
            if (reader.Entry.Key == "[Content_Types].xml")
            {
                using EntryStream xmlStream = reader.OpenEntryStream();
                XDocument xmlDoc = XDocument.Load(xmlStream);
                IEnumerable<XElement> typesElements = xmlDoc.Elements().First().Elements();
                IEnumerable<XElement> notesMasters = typesElements.Where(e =>
                    e.Name.LocalName == "Override" && e.Attribute("ContentType").Value == NotesMasterType);
                if (notesMasters.Count() > 1)
                {
                    return new ValidateResponse("Presentation has more than one Notes Master Part");
                }
            }
        }

        return new ValidateResponse();
    }
}

public class ValidateResponse
{
    public bool IsValid { get; }
    public string ErrorMessage { get; }

    public ValidateResponse()
    {
        IsValid = true;
    }

    public ValidateResponse(string errorMessage)
    {
        ErrorMessage = errorMessage;
    }
}