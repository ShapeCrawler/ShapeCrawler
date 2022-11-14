# Changelog

## Version 0.37.0 - 2022-11-14
- Added `IPicture.SvgContent` property to read SVG graphic content [#344](https://github.com/ShapeCrawler/ShapeCrawler/issues/355)
- Added `ITextFrame.LeftMargin`, `ITextFrame.RightMargin`, `ITextFrame.TopMargin` and `ITextFrame.BottomMargin` properties to get margins of text box [#375](https://github.com/ShapeCrawler/ShapeCrawler/issues/375)
- Added `IParagraph.IndentLevel` to get indent level of paragraph [#377](https://github.com/ShapeCrawler/ShapeCrawler/issues/377)

## Version 0.36.0 - 2022-10-31
- Added `IShapeFill.SetHexSolidColor(string hex)` to set a solid color as the shape fill [#267](https://github.com/ShapeCrawler/ShapeCrawler/issues/267)

## Version 0.35.0 - 2022-10-17
- Added `IShapeFill.SetPicture(Stream image)` to set picture fill [#279](https://github.com/ShapeCrawler/ShapeCrawler/issues/279)
- Added `IFont.OffsetEffect` property to add superscript [#258](https://github.com/ShapeCrawler/ShapeCrawler/issues/258)

## Version 0.34.1 - 2022-10-02
- Fixed updating text of text frame [#332](https://github.com/ShapeCrawler/ShapeCrawler/issues/332)

## Version 0.34.0 - 2022-09-30
- Fixed updating text of Subtitle [#325](https://github.com/ShapeCrawler/ShapeCrawler/issues/325)
- Added `ITableRow.Clone()` to create a row duplication [#326](https://github.com/ShapeCrawler/ShapeCrawler/issues/326)

## Version 0.33.0 - 2022-09-23
- Added `IParagraph.AddPortion(string text)` to add a new text portion in paragraph [#297](https://github.com/ShapeCrawler/ShapeCrawler/issues/297).
- Added APIs to update Underline, Type, Character, Size and Font of paragraph bullet [#311](https://github.com/ShapeCrawler/ShapeCrawler/issues/311).
- Fixed incorrect updating grouped Picture [#295](https://github.com/ShapeCrawler/ShapeCrawler/issues/295).

## Version 0.32.0 - 2022-09-09
- Added opportunity to update text of master shape [#37](https://github.com/ShapeCrawler/ShapeCrawler/issues/37).
- Added `IColorFormat.SetColorHex()` to update color [#37](https://github.com/ShapeCrawler/ShapeCrawler/issues/37).
- Added `IAudioShape.MIME` and `IVideoShape.MIME` to get MIME type of audio and video content [#284](https://github.com/ShapeCrawler/ShapeCrawler/issues/284).

## Version 0.31.2 - 2022-09-01
- Fixed getting binary content of audio and video shapes [#268](https://github.com/ShapeCrawler/ShapeCrawler/issues/268).

## Version 0.31.1 - 2022-07-15
- Fixed bug in Chart [#259](https://github.com/ShapeCrawler/ShapeCrawler/issues/259).

## Version 0.31.0 - 2022-06-10
- Added opportunity to update series value eg. `chart.SeriesCollection[0].Points[0].Value = 10` [#66](https://github.com/ShapeCrawler/ShapeCrawler/issues/66).
- Fixed section slide removing [#240](https://github.com/ShapeCrawler/ShapeCrawler/issues/240).

## Version 0.30.0 - 2022-05-23
- Added `IPresentation.Sections` to access presentation sections [#240](https://github.com/ShapeCrawler/ShapeCrawler/issues/240).
- Fixed issue when `IPresentation.SaveAs()` modifies original presentation [#237](https://github.com/ShapeCrawler/ShapeCrawler/issues/237).

## Version 0.29.0 - 2022-05-09
- Added `Image.MIME` property to get image format [#233](https://github.com/ShapeCrawler/ShapeCrawler/issues/233)
- Added `IPortion.Hyperlink` property to add hyperlink [#242](https://github.com/ShapeCrawler/ShapeCrawler/issues/242)

## Version 0.28.1 - 2022-03-21
- Fixed reading picture of Layout and Master slides.

## Version 0.28.0 - 2022-02-10
- Added `IParagraph.Alignment` property for paragraph content alignment.

## Version 0.27.0 - 2022-02-03
- Added support for Connection shape which presents Lines.

## Version 0.26.0 - 2022-01-03
- Added "Shring text on overflow" support for `ITextBox.Text`.

## Version 0.25.0 - 2021-12-16
- Added `IShapeCollection.AddNewVideo()` to add a new video shape on a slide.

## Version 0.24.0 - 2021-09-26
- Added `IShapeCollection.AddNewAudio(int xPixel, int yPixels, Stream mp3Stream)` to add a new audio shape on a slide.
- Added setter for `IShape.Width` and `IShape.Height` properties to change width and height sizes.

## Version 0.23.0 - 2021-09-11
- Added `ISlideCollection.Insert(int position, ISlide outerSlide)` to insert slide at certain position.
- Fixed case when `ISlideCollection.Add()` breaks presentation.

## Version 0.22.0 - 2021-08-14
- Added ability to update chart category.

## Version 0.21.1 - 2021-07-30
### Fixed
- Fixed `IPresentation.SaveAs()`. It did not release underlying resources in the right way.

## Version 0.21.0 - 2021-06-23
### Added
- Added `void ISlideCollection.Add(ISlide addingSlide)` to add outer slide.
- Added setter for `ISlide.Number` to change slide position.

## Version 0.20.1 - 2021-06-07
### Fixed
- Fixed changing picture source with shared image source.

## Version 0.20.0 - 2021-05-08
### Added
- Added `Portion.Font.ColorFormat` to read color properties of font.

## Version 0.19.0 - 2021-04-13
### Added
- Added .NET Standard 2.0 target.

## Version 0.18.0 - 2021-03-28
### Added
- Added setter for `IFont.IsBold` property to set up bold font.
- Added `IFont.IsItalic` property to define whether font is italic.

## Version 0.17.0 - 2021-03-21
### Added
- Added `IFont.IsBold` property to define whether font is bold.

## Version 0.16.1 - 2021-03-08
### Fixed
- Fixed parser of font properties

## Version 0.16.0 - 2021-02-20
### Added
- Added `ITable.MergeCells()` API to merge neigbor cells of the table (#109)

## Version 0.15.0 - 2021-02-13
### Added
- Added setter for `Column.Width` to change width of a table column (#105) 
- Added `Row.Height` property to access height of table row (#105)

## Version 0.14.0 - 2021-01-31
### Added
- Added two-dimensional indexer for `TableSc[int row_index][int column_Index]` to get table cell by row and column indexes (#29)
- Added support for .NET 5 (#98)
- Added `Column.Width` to get width of table column (#101)

## Version 0.13.0 - 2021-01-24
### Added
- Added `CellSc.IsMergedCell` to define whether table cell belong to merged cells group (#35)
- Added `ParagraphCollection.Add()` method to add a new paragraph (#62)

## Version 0.12.0 - 2021-01-17
### Added
- Added base API to get Slide Master collection — `PresentationSc.SlideMasters` (#44)
### Fixed
- Fixed New Line character processing for text paragraph (#87)
## Version 0.11.0 - 2021-01-10
### Added
- Added setter for `Portion.Text` property to be able to change text of paragraph portion (#22)
- Added setter for `Portion.Font.Name` to change font name of the portion of non-placeholder shape (#82)
- Added setter for `Portion.Font.Size` to change font size of the portion of non-placeholer shape (#81)
## Version 0.10.0 - 2021-01-01
### Added
- Added `Portion.Remove()` to be able to remove paragraph portion;
- Added setter for `Paragraph.Text` property to be able to change paragraph's text;
- Added support for .NET Core 2.0

## Version 0.9.0 - 2020-12-24
### Added
- Added `Slide.Hide()` and `Slide.Hidden` APIs to hide slide and define whether the slide is hidden;
- Added support .NET Standard 2.0 and .NET Standard 2.1 frameworks.

## Version 0.8.0 - 2020-12-20
### Added
- Added `CustomData` property for slide and shape objects: `Slide.CustomData`, `ShapeEx.CustomData`. These property allows to store some user's custom string.

## Version 0.7.0 - 2020-10-12
### Added
- Added `Bullet` property for the paragraph:
    - Bullet.Type
    - Bullet.Char
    - Bullet.FontName
    - Bullet.Size
    - Bullet.ColorHex

## Version 0.6.0 - 2020-05-31
### Added
- Added `Series.Name` property;
- Added `SlideEx.SaveScheme()` to save slide's scheme to PNG file:
![slide-scheme](/resources/slide-scheme.png)

## Version 0.5.0 - 2020-05-02
### Added
- Added `ShapeEx.GeometryType` property contaning the geometric form:
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
- Added `ChartEx.XValues` property for charts like ScatterChart.

## Version 0.4.0 - 2020-03-28
### Added
- Added setters for `X`, `Y`, `Width` and `Height` properties of non-placeholder shapes;
- Added `ShapeEx.IsGrouped` boolean property to determine whether the shape is grouped;
- Added APIs to remove table row
  ```
  var tableRows = shape.Table.Rows;
  // remove by index
  tableRows.RemoveAt(0);
  // remove by instance
  var row = tableRows.First();
  tableRows.Remove(row);
  ```

## Version 0.3.0 - 2020-03-16
### Added
- Added _ChartEx.SeriesCollection_ and  _ChartEx.Categories_ collections
    ```
    var pointValue = chart.SeriesCollection[0].PointValues[0];
    var seriesType = chart.SeriesCollection[0].Type;
    if (chart.HasCategories)
    {
        var category = chart.Categories[0];
    }
    ```

## Version 0.2.0 - 2020-03-02
### Added
- Added `SlideNumber` placeholder processing;
- Added property `ShapeEx.Fill`.

## Version 0.1.0 - 2020-02-25
### Added
- Initial release of SlideDotNet.
