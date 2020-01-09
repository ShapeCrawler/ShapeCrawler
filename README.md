# PptxXML
PptxXML is a lightweight .NET library for parse PowerPoint file presentations. It aims to provide an intuitive and user-friendly interface to dealing with the underlying [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) API.

When working with PowerPoint documents, developers traditionally choose to either use raw XML or rely on the Office Automation libraries. But as most of you know, the Office Automation library is not appropriate for servers and working with XML can be quite tedious. PptxXML bridges the gap by providing an easy to use API without the overhead of COM.

# Install PptxXML via NuGet
To install PptxXML, run the following command in the Package Manager Console
`PM> Install-Package PptxXML`

# What can you do with this?
PptxXML allows you to parse PowerPoint files without the PowerPoint application. The typical example is processing PowerPoint files on a web server.

**Example #1** (count number of slides):
...

**Example #1** (remove slide):

# If you have problems
If you have "how-to" questions please post [Stack Overflow](https://stackoverflow.com/) with **pptx-xml** tag.

# Author