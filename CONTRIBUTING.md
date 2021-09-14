# Recomended tools
The internal structure of PowerPoint presentation is one the most difficult among other documents from the Microsoft Office suite. For example, the slide presented for a user is not a single document but only top on Slide Layout and Slide Master slides. Even just that levels frequently lead to confusion. The following is a list of tools and notes that should simplify your development while working on ShapeCrawler's issue.

- **[OOXML Viewer](https://marketplace.visualstudio.com/items?itemName=yuenm18.ooxml-viewer)** — extension for Visual Studio Code allowing to view Open XML package. One of the good features of this extension is track changes of modified presentation.
- **[Open XML SDK 2.5 Productivity Tool](https://github.com/OfficeDev/Open-XML-SDK/releases/tag/v2.5)** — application for generating C#-code from Open XML document. It can be useful, for example, when you wanna know what C#-code is needed to add a new shape or slide.
- **.pptx is ZIP** — do not forget that .pptx-file is a zip file as well as other Open XML documents. Thus you can rename the extension of the presentation file on .zip and watch his internals:

![.pptx is zip](/resources/pptx_is_zip.gif)