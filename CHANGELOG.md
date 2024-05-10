# Changelog  

## Version 0.50.4 - 2024-05-10
🐞Fixed `ISlideShapes.AddPicture()` [#671](https://github.com/ShapeCrawler/ShapeCrawler/issues/671)

## Version 0.50.3 - 2024-03-06
🐞Fixed `IShape.AsTable()`

## Version 0.50.2 - 2024-03-04
🐞Fixed slide adding

## Version 0.50.1 - 2023-12-08
🍀Added `IShape.SDKPath` to store the XPath of the underlying Open XML element [#592](https://github.com/ShapeCrawler/ShapeCrawler/issues/592)  

## Version 0.49.0 - 2023-09-12
🍀Added new `SCAudioType` to be able to add audio shape with different types [#579](https://github.com/ShapeCrawler/ShapeCrawler/issues/579)  
🐞Fixed an issue with Slide Background updating [#577](https://github.com/ShapeCrawler/ShapeCrawler/issues/577)  

## Version 0.48.0 - 2023-08-19
🍀Added new properties: `IShapeFill.AlphaPercentage`, `IShapeFill.LuminanceModulationPercentage` and `IShapeFill.LuminanceOffsetPercentage` to the in shape filling object [#567](https://github.com/ShapeCrawler/ShapeCrawler/issues/567)   
🍀Added a new property: `Shape.Rotation` to the shape object  
🐞Fixed an issue with Shape Fill [#558](https://github.com/ShapeCrawler/ShapeCrawler/issues/558)  
🐞Fixed merging of table cells [#564](https://github.com/ShapeCrawler/ShapeCrawler/issues/564)

## Version 0.47.0 - 2023-07-26
🍀Added setters for `IParagraph.IndentLevel`  
🍀Added `IParagraph.HeaderAndFooter.AddSlideNumber()` [#540](https://github.com/ShapeCrawler/ShapeCrawler/issues/540)

## Version 0.46.0 - 2023-07-07
🍀Added setters for `IPresentation.SlideHeight/SlideWidth` [#522](https://github.com/ShapeCrawler/ShapeCrawler/issues/522)  
🍀Added `IShapeCollection.Add()` [#264](https://github.com/ShapeCrawler/ShapeCrawler/issues/264)  
🐞Fixed `ISlide.Number` setter  
🐞Fixed updating height of table row [#532](https://github.com/ShapeCrawler/ShapeCrawler/issues/532)

## Version 0.45.3 - 2023-06-24
🐞Fixed updating Hyperlink [#518](https://github.com/ShapeCrawler/ShapeCrawler/issues/518)

## Version 0.45.2 - 2023-06-18
🐞Fixed bug in `IPresentation.BinaryData` [#515](https://github.com/ShapeCrawler/ShapeCrawler/issues/515)

## Version 0.45.1 - 2023-05-18
🐞Fixed bug where `ISlideCollection.Add()` doesn't copy placeholder shapes [#508](https://github.com/ShapeCrawler/ShapeCrawler/issues/508)

## Version 0.45.0 - 2023-05-05
🍀Added setters for `IChart.Axes.ValueAxis.Minumum/Maximum`[#482](https://github.com/ShapeCrawler/ShapeCrawler/issues/482)  
🍀Added `ISeriesCollection.RemoveAt(int index)` [#491](https://github.com/ShapeCrawler/ShapeCrawler/issues/491)  
🍀Added `ITable.RemoveColumnAt(int columnIndex)` [#501](https://github.com/ShapeCrawler/ShapeCrawler/issues/501)  
🐞Fixed updating text of the grouped shape [#452](https://github.com/ShapeCrawler/ShapeCrawler/issues/452)  

## Version 0.44.0 - 2023-04-21
🍀Added `IShapeCollection.AddPicture()` [#481](https://github.com/ShapeCrawler/ShapeCrawler/issues/481)   
🍀Added `IChart.FormatAxis.AxisOptions.Bounds.Minimum/Maximum` [#482](https://github.com/ShapeCrawler/ShapeCrawler/issues/482)  

## Version 0.43.0 - 2023-04-06
🍀Added `IShapeCollection.AddLine()` to add Line shape [#465](https://github.com/ShapeCrawler/ShapeCrawler/issues/465)

## Version 0.42.1 - 2023-03-17
🐞Fixed the table cell merging problem [#472](https://github.com/ShapeCrawler/ShapeCrawler/issues/472)  
🐞Fixed text alignment [#476](https://github.com/ShapeCrawler/ShapeCrawler/issues/476)

## Version 0.42.0 - 2023-03-04
🍀Added `IAutoShape.Duplicate()` [#444](https://github.com/ShapeCrawler/ShapeCrawler/issues/444)  
🍀Added `IShapeCollection.AddLine()` [#465](https://github.com/ShapeCrawler/ShapeCrawler/issues/465)

## Version 0.41.4 - 2023-02-13
🐞Fixed updating X/Y coordinates of grouped shape [#452](https://github.com/ShapeCrawler/ShapeCrawler/issues/452)

## Version 0.41.3 - 2023-01-29
🐞Fixed solid color setting [d442](https://github.com/ShapeCrawler/ShapeCrawler/discussions/442)

## Version 0.41.2 - 2023-01-28
🐞Fixed updating Table coordinates [d443](https://github.com/ShapeCrawler/ShapeCrawler/discussions/443)

## Version 0.41.1 - 2022-01-13
🐞Fixed East Asian font parsing.  
🐞Fixed adding a new shape.

## Version 0.41.0 - 2022-01-10
🍀Added supporting East Asian fonts [#419](https://github.com/ShapeCrawler/ShapeCrawler/issues/419)  
🍀Added `IAutoShapeCollection.AddRoundedRectangle()`  
  
## Version 0.40.0 - 2022-12-26  
🍀Added `ISlideCollection.AddEmptySlide()` [#141](https://github.com/ShapeCrawler/ShapeCrawler/issues/141)    
🍀Added `IShapeCollection.Remove()` [#34](https://github.com/ShapeCrawler/ShapeCrawler/issues/34)    
🍀Added `ISlideMaster.ITheme` [#369](https://github.com/ShapeCrawler/ShapeCrawler/issues/369)    
  
## Version 0.39.0 - 2022-12-12  
🍀Added setter for `ITextFrame.LeftMargin`, `ITextFrame.RightMargin`, `ITextFrame.TopMargin` and `ITextFrame.BottomMargin` properties [#385](https://github.com/ShapeCrawler/ShapeCrawler/issues/385)    
🍀Added `IPortion.TextHighlightColor` [#139](https://github.com/ShapeCrawler/ShapeCrawler/issues/139)    
🍀Added `IParagraph.Spacing` [#379](https://github.com/ShapeCrawler/ShapeCrawler/issues/379)    
🍀Added `IAutoShape.Outline` [#373](https://github.com/ShapeCrawler/ShapeCrawler/issues/373)    
🍀Added `IShapeCollection.AddAutoShape()` [#53](https://github.com/ShapeCrawler/ShapeCrawler/issues/53)    
🍀Added `IShapeCollection.AddTable()` [#53](https://github.com/ShapeCrawler/ShapeCrawler/issues/53)  
🍀Added `IRowCollection.Add()` [#309](https://github.com/ShapeCrawler/ShapeCrawler/issues/309)  
  
## Version 0.38.0 - 2022-11-28  
🍀Added setter for `ITextFrame.AutofitType` property [#360](https://github.com/ShapeCrawler/ShapeCrawler/issues/360)  
  
## Version 0.37.1 - 2022-11-24  
🐞Fixed `IPortion.Hyperlink` [#394](https://github.com/ShapeCrawler/ShapeCrawler/discussions/394)  
  
## Version 0.37.0 - 2022-11-14  
🍀Added `IPicture.SvgContent` property to read SVG graphic content [#344](https://github.com/ShapeCrawler/ShapeCrawler/issues/355)  
🍀Added `ITextFrame.LeftMargin`, `ITextFrame.RightMargin`, `ITextFrame.TopMargin` and `ITextFrame.BottomMargin` properties to get margins of text box [#375](https://github.com/ShapeCrawler/ShapeCrawler/issues/375)  
🍀Added `IParagraph.IndentLevel` to get indent level of paragraph [#377](https://github.com/ShapeCrawler/ShapeCrawler/issues/377)  
  
## Version 0.36.0 - 2022-10-31  
🍀Added `IShapeFill.SetHexSolidColor(string hex)` to set a solid color as the shape fill [#267](https://github.com/ShapeCrawler/ShapeCrawler/issues/267)  
  
## Version 0.35.0 - 2022-10-17  
🍀Added `IShapeFill.SetPicture(Stream image)` to set picture fill [#279](https://github.com/ShapeCrawler/ShapeCrawler/issues/279)  
🍀Added `IFont.OffsetEffect` property to add superscript [#258](https://github.com/ShapeCrawler/ShapeCrawler/issues/258)  
  
## Version 0.34.1 - 2022-10-02  
🐞Fixed updating text of text frame [#332](https://github.com/ShapeCrawler/ShapeCrawler/issues/332)  
  
## Version 0.34.0 - 2022-09-30  
🐞Fixed updating text of Subtitle [#325](https://github.com/ShapeCrawler/ShapeCrawler/issues/325)  
🍀Added `ITableRow.Clone()` to create a row duplication [#326](https://github.com/ShapeCrawler/ShapeCrawler/issues/326)  
  
## Version 0.33.0 - 2022-09-23  
🍀Added `IParagraph.AddPortion(string text)` to add a new text portion in paragraph [#297](https://github.com/ShapeCrawler/ShapeCrawler/issues/297).  
🍀Added APIs to update Underline, Type, Character, Size and Font of paragraph bullet [#311](https://github.com/ShapeCrawler/ShapeCrawler/issues/311).  
🐞Fixed incorrect updating grouped Picture [#295](https://github.com/ShapeCrawler/ShapeCrawler/issues/295).  
  
## Version 0.32.0 - 2022-09-09  
🍀Added opportunity to update text of master shape [#37](https://github.com/ShapeCrawler/ShapeCrawler/issues/37).  
🍀Added `IColorFormat.SetColorHex()` to update color [#37](https://github.com/ShapeCrawler/ShapeCrawler/issues/37).  
🍀Added `IAudioShape.MIME` and `IVideoShape.MIME` to get MIME type of audio and video content [#284](https://github.com/ShapeCrawler/ShapeCrawler/issues/284).  
  
## Version 0.31.2 - 2022-09-01  
🐞Fixed getting binary content of audio and video shapes [#268](https://github.com/ShapeCrawler/ShapeCrawler/issues/268).  
  
## Version 0.31.1 - 2022-07-15  
🐞Fixed bug in Chart [#259](https://github.com/ShapeCrawler/ShapeCrawler/issues/259).  
  
## Version 0.31.0 - 2022-06-10  
🍀Added opportunity to update series value eg. `chart.SeriesCollection[0].Points[0].Value = 10` [#66](https://github.com/ShapeCrawler/ShapeCrawler/issues/66).  
🐞Fixed section slide removing [#240](https://github.com/ShapeCrawler/ShapeCrawler/issues/240).  
  
## Version 0.30.0 - 2022-05-23  
🍀Added `IPresentation.Sections` to access presentation sections [#240](https://github.com/ShapeCrawler/ShapeCrawler/issues/240).  
🐞Fixed issue when `IPresentation.SaveAs()` modifies original presentation [#237](https://github.com/ShapeCrawler/ShapeCrawler/issues/237).  
  
## Version 0.29.0 - 2022-05-09  
🍀Added `Image.MIME` property to get image format [#233](https://github.com/ShapeCrawler/ShapeCrawler/issues/233)  
🍀Added `IPortion.Hyperlink` property to add hyperlink [#242](https://github.com/ShapeCrawler/ShapeCrawler/issues/242)  
  
## Version 0.28.1 - 2022-03-21  
🐞Fixed reading picture of Layout and Master slides.  
  
## Version 0.28.0 - 2022-02-10  
🍀Added `IParagraph.Alignment` property for paragraph content alignment.  
  
## Version 0.27.0 - 2022-02-03  
🍀Added support for Connection shape which presents Lines.  
  
## Version 0.26.0 - 2022-01-03  
🍀Added "Shring text on overflow" support for `ITextBox.Text`.  
  
## Version 0.25.0 - 2021-12-16  
🍀Added `IShapeCollection.AddNewVideo()` to add a new video shape on a slide.  
  
## Version 0.24.0 - 2021-09-26  
🍀Added `IShapeCollection.AddNewAudio(int xPixel, int yPixels, Stream mp3Stream)` to add a new audio shape on a slide.  
🍀Added setter for `IShape.Width` and `IShape.Height` properties to change width and height sizes.  
  
## Version 0.23.0 - 2021-09-11  
🍀Added `ISlideCollection.Insert(int position, ISlide outerSlide)` to insert slide at certain position.  
🐞Fixed case when `ISlideCollection.Add()` breaks presentation.  
  
## Version 0.22.0 - 2021-08-14  
🍀Added ability to update chart category.  
  
## Version 0.21.1 - 2021-07-30  
🐞Fixed `IPresentation.SaveAs()`. It did not release underlying resources in the right way.  
  
## Version 0.21.0 - 2021-06-23  
  
🍀Added `void ISlideCollection.Add(ISlide addingSlide)` to add outer slide.  
🍀Added setter for `ISlide.Number` to change slide position.  
  
## Version 0.20.1 - 2021-06-07  
  
🐞Fixed changing picture source with shared image source.  
  
## Version 0.20.0 - 2021-05-08  
  
🍀Added `Portion.Font.ColorFormat` to read color properties of font.  
  
## Version 0.19.0 - 2021-04-13  

🍀Added .NET Standard 2.0 target.  
  
## Version 0.18.0 - 2021-03-28  
  
🍀Added setter for `IFont.IsBold` property to set up bold font.  
🍀Added `IFont.IsItalic` property to define whether font is italic.  
  
## Version 0.17.0 - 2021-03-21  
  
🍀Added `IFont.IsBold` property to define whether font is bold.  
  
## Version 0.16.1 - 2021-03-08  
  
🐞Fixed parser of font properties  
  
## Version 0.16.0 - 2021-02-20  
  
🍀Added `ITable.MergeCells()` API to merge neigbor cells of the table (#109)  
  
## Version 0.15.0 - 2021-02-13  
  
🍀Added setter for `Column.Width` to change width of a table column (#105)   
🍀Added `Row.Height` property to access height of table row (#105)  
  
## Version 0.14.0 - 2021-01-31  
  
🍀Added two-dimensional indexer for `TableSc[int row_index][int column_Index]` to get table cell by row and column indexes (#29)  
🍀Added support for .NET 5 (#98)  
🍀Added `Column.Width` to get width of table column (#101)  
  
## Version 0.13.0 - 2021-01-24  
  
🍀Added `CellSc.IsMergedCell` to define whether table cell belong to merged cells group (#35)  
🍀Added `ParagraphCollection.Add()` method to add a new paragraph (#62)  
  
## Version 0.12.0 - 2021-01-17  
  
🍀Added base API to get Slide Master collection — `PresentationSc.SlideMasters` (#44)    
🐞Fixed New Line character processing for text paragraph (#87)  
## Version 0.11.0 - 2021-01-10  
  
🍀Added setter for `Portion.Text` property to be able to change text of paragraph portion (#22)  
🍀Added setter for `Portion.Font.Name` to change font name of the portion of non-placeholder shape (#82)  
🍀Added setter for `Portion.Font.Size` to change font size of the portion of non-placeholer shape (#81)  
## Version 0.10.0 - 2021-01-01  
  
🍀Added `Portion.Remove()` to be able to remove paragraph portion;  
🍀Added setter for `Paragraph.Text` property to be able to change paragraph's text;  
🍀Added support for .NET Core 2.0  
  
## Version 0.9.0 - 2020-12-24  
  
🍀Added `Slide.Hide()` and `Slide.Hidden` APIs to hide slide and define whether the slide is hidden;  
🍀Added support .NET Standard 2.0 and .NET Standard 2.1 frameworks.  
  
## Version 0.8.0 - 2020-12-20  
  
🍀Added `CustomData` property for slide and shape objects: `Slide.CustomData`, `ShapeEx.CustomData`. These property allows to store some user's custom string.  
  
## Version 0.7.0 - 2020-10-12  
  
🍀Added `Bullet` property for the paragraph:  
    - Bullet.Type  
    - Bullet.Char  
    - Bullet.FontName  
    - Bullet.Size  
    - Bullet.ColorHex  
  
## Version 0.6.0 - 2020-05-31  
  
🍀Added `Series.Name` property  
🍀Added `SlideEx.SaveScheme()` to save slide's scheme to PNG file
  
## Version 0.5.0 - 2020-05-02  
  
🍀Added `ShapeEx.GeometryType` property contaning the geometric form:  
```  
public enum GeometryType  
{  
    Line,  
    LineInverse,  
    Triangle,  
    RightTriangle,  
    Rectangle,  
    ...  
```  
🍀Added `ChartEx.XValues` property for charts like ScatterChart.  
  
## Version 0.4.0 - 2020-03-28  
  
🍀Added setters for `X`, `Y`, `Width` and `Height` properties of non-placeholder shapes;  
🍀Added `ShapeEx.IsGrouped` boolean property to determine whether the shape is grouped;  
🍀Added APIs to remove table row  
  ```  
  var tableRows = shape.Table.Rows;  
  // remove by index  
  tableRows.RemoveAt(0);  
  // remove by instance  
  var row = tableRows.First();  
  tableRows.Remove(row);  
  ```  
  
## Version 0.3.0 - 2020-03-16  
  
🍀Added _ChartEx.SeriesCollection_ and  _ChartEx.Categories_ collections  
    ```  
    var pointValue = chart.SeriesCollection[0].PointValues[0];  
    var seriesType = chart.SeriesCollection[0].Type;  
    if (chart.HasCategories)  
    {  
        var category = chart.Categories[0];  
    }  
    ```  
  
## Version 0.2.0 - 2020-03-02  
  
🍀Added `SlideNumber` placeholder processing;  
🍀Added property `ShapeEx.Fill`.  
  
## Version 0.1.0 - 2020-02-25  
  
- Initial release of SlideDotNet.  
  