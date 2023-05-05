using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using C14 = DocumentFormat.OpenXml.Office2010.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Factories;

internal sealed class ChartGraphicFrameHandler : OpenXmlElementHandler
{
    private const string Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public P.GraphicFrame Create(TypedOpenXmlPart typedOpenXmlPart)
    {
        var id = (UInt32Value)6U;
        var name = "Chart X";

        var chartPart = typedOpenXmlPart.AddNewPart<ChartPart>("rId2");
        this.GenerateChartPartContent(chartPart);

        // Create Excel
        var embeddedPackagePart = chartPart.AddNewPart<EmbeddedPackagePart>(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
        var embeddedPackagePartStream = embeddedPackagePart.GetStream(FileMode.Create);
        using var spreadsheetDocument = SpreadsheetDocument.Create(
            embeddedPackagePartStream,
            SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new X.Workbook();
        var sheets = new X.Sheets();
        workbookPart.Workbook.AppendChild(sheets);

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new X.SheetData();
        worksheetPart.Worksheet = new X.Worksheet(sheetData);
        var sheet = new X.Sheet
        {
            Id = spreadsheetDocument.WorkbookPart!.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        };
        sheets.Append(sheet);

        var cellB1 = new X.Cell
            { CellReference = "B1", DataType = X.CellValues.String, CellValue = new X.CellValue("Series 1") };
        var cellC1 = new X.Cell
            { CellReference = "C1", DataType = X.CellValues.String, CellValue = new X.CellValue("Series 2") };
        var row1 = new X.Row { RowIndex = 1 };
        row1.Append(cellB1);
        row1.Append(cellC1);
        sheetData.Append(row1);

        var cellA2 = new X.Cell
            { CellReference = "A2", DataType = X.CellValues.String, CellValue = new X.CellValue("Category 1") };
        var cellB2 = new X.Cell
            { CellReference = "B2", DataType = X.CellValues.Number, CellValue = new X.CellValue("1") };
        var cellC2 = new X.Cell
            { CellReference = "C2", DataType = X.CellValues.Number, CellValue = new X.CellValue("2") };
        var row2 = new X.Row { RowIndex = 2 };
        row2.Append(cellA2);
        row2.Append(cellB2);
        row2.Append(cellC2);
        sheetData.Append(row2);

        var cellA3 = new X.Cell
            { CellReference = "A3", DataType = X.CellValues.String, CellValue = new X.CellValue("Category 2") };
        var cellB3 = new X.Cell
            { CellReference = "B3", DataType = X.CellValues.Number, CellValue = new X.CellValue("3") };
        var cellC3 = new X.Cell
            { CellReference = "C3", DataType = X.CellValues.Number, CellValue = new X.CellValue("4") };
        var row3 = new X.Row { RowIndex = 3 };
        row3.Append(cellA3);
        row3.Append(cellB3);
        row3.Append(cellC3);
        sheetData.Append(row3);

        var cellA4 = new X.Cell
            { CellReference = "A4", DataType = X.CellValues.String, CellValue = new X.CellValue("Category 3") };
        var cellB4 = new X.Cell
            { CellReference = "B4", DataType = X.CellValues.Number, CellValue = new X.CellValue("5") };
        var cellC4 = new X.Cell
            { CellReference = "C4", DataType = X.CellValues.Number, CellValue = new X.CellValue("6") };
        var row4 = new X.Row { RowIndex = 4 };
        row4.Append(cellA4);
        row4.Append(cellB4);
        row4.Append(cellC4);
        sheetData.Append(row4);

        spreadsheetDocument.Save();
        spreadsheetDocument.Dispose();
        embeddedPackagePartStream.Dispose();

        // Create Shape
        var graphicFrame = new P.GraphicFrame();

        var nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = id, Name = name };

        var nonVisualDrawingPropertiesExtensionList = new A.NonVisualDrawingPropertiesExtensionList();

        var nonVisualDrawingPropertiesExtension1 =
            new A.NonVisualDrawingPropertiesExtension { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties.Append(nonVisualDrawingPropertiesExtensionList);
        var nonVisualGraphicFrameDrawingProperties =
            new P.NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nonVisualGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);

        var transform = new P.Transform();
        var offset = new A.Offset { X = 2032000L, Y = 719666L };
        var extents = new A.Extents { Cx = 8128000L, Cy = 5418667L };

        transform.Append(offset);
        transform.Append(extents);

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

        var chartReference = new C.ChartReference { Id = "rId2" };
        chartReference.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        graphicData.Append(chartReference);

        graphic.Append(graphicData);

        graphicFrame.Append(nonVisualGraphicFrameProperties);
        graphicFrame.Append(transform);
        graphicFrame.Append(graphic);

        return graphicFrame;
    }

    internal override SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideStructure,
        OneOf<ShapeCollection, SCGroupShape> shapeCollection)
    {
        if (pShapeTreeChild is not P.GraphicFrame pGraphicFrame)
        {
            return this.Successor?.FromTreeChild(pShapeTreeChild, slideStructure, shapeCollection);
        }

        var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
        if (!aGraphicData.Uri!.Value!.Equals(Uri, StringComparison.Ordinal))
        {
            return this.Successor?.FromTreeChild(pShapeTreeChild, slideStructure, shapeCollection);
        }

        var slideBase = slideStructure.Match(slide => slide as SlideStructure, layout => layout, master => master);
        var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
        var chartPart = (ChartPart)slideBase.TypedOpenXmlPart.GetPartById(cChartRef.Id!);
        var cPlotArea = chartPart!.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
        var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        if (cCharts.Count() > 1)
        {
            return new SCComboChart(pGraphicFrame, slideStructure, shapeCollection);
        }

        var chartTypeName = cCharts.Single().LocalName;

        if (chartTypeName == "lineChart")
        {
            return new SCLineChart(pGraphicFrame, slideStructure, shapeCollection);
        }

        if (chartTypeName == "barChart")
        {
            return new SCBarChart(pGraphicFrame, slideStructure, shapeCollection);
        }

        if (chartTypeName == "pieChart")
        {
            return new SCPieChart(pGraphicFrame, slideStructure, shapeCollection);
        }

        if (chartTypeName == "scatterChart")
        {
            return new SCScatterChart(pGraphicFrame, slideStructure, shapeCollection);
        }

        return new SCChart(pGraphicFrame, slideStructure, shapeCollection);
    }

    private void GenerateChartPartContent(ChartPart chartPart1)
    {
        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        chartSpace.AddNamespaceDeclaration("c16r2", "http://schemas.microsoft.com/office/drawing/2015/06/chart");
        var date19041 = new C.Date1904 { Val = false };
        var editingLanguage1 = new C.EditingLanguage { Val = "en-US" };
        var roundedCorners1 = new C.RoundedCorners { Val = false };

        AlternateContent alternateContent1 = new AlternateContent();
        alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

        AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "c14" };
        alternateContentChoice1.AddNamespaceDeclaration("c14", "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
        C14.Style style1 = new C14.Style() { Val = 102 };

        alternateContentChoice1.Append(style1);

        AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();
        C.Style style2 = new C.Style() { Val = 2 };

        alternateContentFallback1.Append(style2);

        alternateContent1.Append(alternateContentChoice1);
        alternateContent1.Append(alternateContentFallback1);

        var cChart = new C.Chart();

        var title = new C.Title();
        var overlay = new C.Overlay() { Val = false };

        var chartShapeProperties = new C.ChartShapeProperties();
        var noFill1 = new A.NoFill();

        var outline1 = new A.Outline();
        var noFill2 = new A.NoFill();

        outline1.Append(noFill2);
        A.EffectList effectList1 = new A.EffectList();

        chartShapeProperties.Append(noFill1);
        chartShapeProperties.Append(outline1);
        chartShapeProperties.Append(effectList1);

        C.TextProperties textProperties1 = new C.TextProperties();
        A.BodyProperties bodyProperties1 = new A.BodyProperties()
        {
            Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true
        };
        A.ListStyle listStyle1 = new A.ListStyle();

        A.Paragraph paragraph1 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties()
        {
            FontSize = 1862, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0
        };

        A.SolidFill solidFill10 = new A.SolidFill();

        A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor10.Append(luminanceModulation1);
        schemeColor10.Append(luminanceOffset1);

        solidFill10.Append(schemeColor10);
        A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties11.Append(solidFill10);
        defaultRunProperties11.Append(latinFont10);
        defaultRunProperties11.Append(eastAsianFont10);
        defaultRunProperties11.Append(complexScriptFont10);

        paragraphProperties1.Append(defaultRunProperties11);
        A.EndParagraphRunProperties endParagraphRunProperties1 =
            new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph1.Append(paragraphProperties1);
        paragraph1.Append(endParagraphRunProperties1);

        textProperties1.Append(bodyProperties1);
        textProperties1.Append(listStyle1);
        textProperties1.Append(paragraph1);

        title.Append(overlay);
        title.Append(chartShapeProperties);
        title.Append(textProperties1);
        C.AutoTitleDeleted autoTitleDeleted1 = new C.AutoTitleDeleted() { Val = false };

        C.PlotArea plotArea1 = new C.PlotArea();
        C.Layout layout1 = new C.Layout();

        C.BarChart barChart = new C.BarChart();
        C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
        C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
        C.VaryColors varyColors1 = new C.VaryColors() { Val = false };


        C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
        C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

        // CATEGORY AXIS
        var catAxisData = new C.CategoryAxisData();
        var catStrRef = new C.StringReference();
        var catFormula = new C.Formula();
        catFormula.Text = "Sheet1!$A$2:$A$4";
        var catStrCache = new C.StringCache();
        var catPointCount = new C.PointCount { Val = (UInt32Value)3U };
        var catStrPoint1 = new C.StringPoint { Index = (UInt32Value)0U };
        var catNumValue1 = new C.NumericValue();
        catNumValue1.Text = "Category 1";
        catStrPoint1.Append(catNumValue1);
        var catStrPoint2 = new C.StringPoint { Index = (UInt32Value)1U };
        var catNumValue2 = new C.NumericValue();
        catNumValue2.Text = "Category 2";
        catStrPoint2.Append(catNumValue2);
        var catStrPoint3 = new C.StringPoint { Index = (UInt32Value)2U };
        var catNumValue = new C.NumericValue();
        catNumValue.Text = "Category 3";
        catStrPoint3.Append(catNumValue);
        catStrCache.Append(catPointCount);
        catStrCache.Append(catStrPoint1);
        catStrCache.Append(catStrPoint2);
        catStrCache.Append(catStrPoint3);
        catStrRef.Append(catFormula);
        catStrRef.Append(catStrCache);
        catAxisData.Append(catStrRef);

        // SERIES B
        var seriesBTitleText = new C.SeriesText();
        var seriesBTitleStrRef = new C.StringReference();
        var seriesBTitleFormula = new C.Formula();
        seriesBTitleFormula.Text = "Sheet1!$B$1";
        seriesBTitleText.Append(seriesBTitleStrRef);
        var seriesBTitleStrCache = new C.StringCache(new C.PointCount { Val = (UInt32Value)1U });
        var seriesBTitlePoint = new C.StringPoint { Index = (UInt32Value)0U };
        var seriesBTitlePointValue = new C.NumericValue();
        seriesBTitlePointValue.Text = "Series 1";
        seriesBTitlePoint.Append(seriesBTitlePointValue);
        seriesBTitleStrCache.Append(seriesBTitlePoint);
        seriesBTitleStrRef.Append(seriesBTitleFormula);
        seriesBTitleStrRef.Append(seriesBTitleStrCache);

        var seriesBChartShapeProperties = new C.ChartShapeProperties();
        var solidFill11 = new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 });
        var effectList2 = new A.EffectList();
        seriesBChartShapeProperties.Append(solidFill11);
        seriesBChartShapeProperties.Append(new A.Outline(new A.NoFill()));
        seriesBChartShapeProperties.Append(effectList2);
        var seriesBInvertIfNegative = new C.InvertIfNegative { Val = false };

        var seriesBValues = new C.Values();
        var seriesBNumRef = new C.NumberReference();
        var seriesBFormula = new C.Formula();
        seriesBFormula.Text = "Sheet1!$B$2:$B$4";

        var seriesBNumCache = new C.NumberingCache();
        var seriesBFormatCode = new C.FormatCode();
        seriesBFormatCode.Text = "General";

        var seriesBPointCount = new C.PointCount { Val = (UInt32Value)4U };

        seriesBNumCache.Append(seriesBFormatCode);
        seriesBNumCache.Append(seriesBPointCount);

        var defaultSeriesValues = new[] { 1, 3, 5 };
        for (var i = 0; i < defaultSeriesValues.Length; i++)
        {
            var seriesBNumPoint = this.CreateCNumericPoint(i, defaultSeriesValues[i]);
            seriesBNumCache.Append(seriesBNumPoint);
        }

        seriesBNumRef.Append(seriesBNumCache);
        seriesBNumRef.Append(seriesBFormula);
        seriesBValues.Append(seriesBNumRef);

        C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();

        C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        barSerExtensionList1.Append(barSerExtension1);

        var barChartSeriesB = new C.BarChartSeries();
        barChartSeriesB.Append(index1);
        barChartSeriesB.Append(order1);
        barChartSeriesB.Append(seriesBTitleText);
        barChartSeriesB.Append(seriesBChartShapeProperties);
        barChartSeriesB.Append(seriesBInvertIfNegative);
        barChartSeriesB.Append(catAxisData);
        barChartSeriesB.Append(seriesBValues);
        barChartSeriesB.Append(barSerExtensionList1);

        var barChartSeriesC = new C.BarChartSeries();
        var index2 = new C.Index { Val = (UInt32Value)1U };
        var order2 = new C.Order { Val = (UInt32Value)1U };

        var seriesCTitleRef = new C.StringReference();
        var seriesCTitleFormula = new C.Formula();
        seriesCTitleFormula.Text = "Sheet1!$C$1";

        C.StringCache stringCache3 = new C.StringCache();
        C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)1U };
        C.StringPoint stringPoint6 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue10 = new C.NumericValue();
        numericValue10.Text = "Series 2";
        stringPoint6.Append(numericValue10);
        stringCache3.Append(pointCount4);
        stringCache3.Append(stringPoint6);

        seriesCTitleRef.Append(seriesCTitleFormula);
        seriesCTitleRef.Append(stringCache3);
        var seriesCTitleText = new C.SeriesText();
        seriesCTitleText.Append(seriesCTitleRef);

        var chartShapeProperties3 = new C.ChartShapeProperties();

        A.SolidFill solidFill12 = new A.SolidFill();
        A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent2 };

        solidFill12.Append(schemeColor12);

        A.Outline outline3 = new A.Outline();
        A.NoFill noFill4 = new A.NoFill();

        outline3.Append(noFill4);
        A.EffectList effectList3 = new A.EffectList();

        chartShapeProperties3.Append(solidFill12);
        chartShapeProperties3.Append(outline3);
        chartShapeProperties3.Append(effectList3);
        C.InvertIfNegative invertIfNegative2 = new C.InvertIfNegative() { Val = false };

        var categoryStringReference4 = new C.StringReference();
        var categoryFormula5 = new C.Formula();
        categoryFormula5.Text = "Sheet1!$A$2:$A$4";

        C.StringCache stringCache4 = new C.StringCache();
        C.PointCount pointCount5 = new C.PointCount() { Val = (UInt32Value)4U };

        C.StringPoint stringPoint7 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue11 = new C.NumericValue();
        numericValue11.Text = "Category 1";

        stringPoint7.Append(numericValue11);

        C.StringPoint stringPoint8 = new C.StringPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue12 = new C.NumericValue();
        numericValue12.Text = "Category 2";

        stringPoint8.Append(numericValue12);

        C.StringPoint stringPoint9 = new C.StringPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue13 = new C.NumericValue();
        numericValue13.Text = "Category 3";

        stringPoint9.Append(numericValue13);

        C.StringPoint stringPoint10 = new C.StringPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue14 = new C.NumericValue();
        numericValue14.Text = "Category 4";

        stringPoint10.Append(numericValue14);

        stringCache4.Append(pointCount5);
        stringCache4.Append(stringPoint7);
        stringCache4.Append(stringPoint8);
        stringCache4.Append(stringPoint9);
        stringCache4.Append(stringPoint10);

        categoryStringReference4.Append(categoryFormula5);
        categoryStringReference4.Append(stringCache4);

        var seriesCValues = new C.Values();
        var seriesCNumRef = new C.NumberReference();
        var seriesBValuesFormula = new C.Formula();
        seriesBValuesFormula.Text = "Sheet1!$C$2:$C$4";

        C.NumberingCache numberingCache2 = new C.NumberingCache();
        C.FormatCode formatCode2 = new C.FormatCode();
        formatCode2.Text = "General";
        C.PointCount pointCount6 = new C.PointCount() { Val = (UInt32Value)4U };

        C.NumericPoint numericPoint5 = new C.NumericPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue15 = new C.NumericValue();
        numericValue15.Text = "2.4";

        numericPoint5.Append(numericValue15);

        C.NumericPoint numericPoint6 = new C.NumericPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue16 = new C.NumericValue();
        numericValue16.Text = "4.4000000000000004";

        numericPoint6.Append(numericValue16);

        C.NumericPoint numericPoint7 = new C.NumericPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue17 = new C.NumericValue();
        numericValue17.Text = "1.8";

        numericPoint7.Append(numericValue17);

        C.NumericPoint numericPoint8 = new C.NumericPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue18 = new C.NumericValue();
        numericValue18.Text = "2.8";

        numericPoint8.Append(numericValue18);

        numberingCache2.Append(formatCode2);
        numberingCache2.Append(pointCount6);
        numberingCache2.Append(numericPoint5);
        numberingCache2.Append(numericPoint6);
        numberingCache2.Append(numericPoint7);
        numberingCache2.Append(numericPoint8);

        seriesCNumRef.Append(seriesBValuesFormula);

        seriesCValues.Append(seriesCNumRef);

        C.BarSerExtensionList barSerExtensionList2 = new C.BarSerExtensionList();

        C.BarSerExtension barSerExtension2 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        barSerExtensionList2.Append(barSerExtension2);

        barChartSeriesC.Append(index2);
        barChartSeriesC.Append(order2);
        barChartSeriesC.Append(seriesCTitleText);
        barChartSeriesC.Append(chartShapeProperties3);
        barChartSeriesC.Append(invertIfNegative2);
        barChartSeriesC.Append(catAxisData.CloneNode(true));
        barChartSeriesC.Append(seriesCValues);
        barChartSeriesC.Append(barSerExtensionList2);

        C.StringReference stringReference5 = new C.StringReference();

        C.StringCache stringCache5 = new C.StringCache();
        C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)1U };

        C.StringPoint stringPoint11 = new C.StringPoint() { Index = (UInt32Value)0U };

        stringCache5.Append(pointCount7);
        stringCache5.Append(stringPoint11);

        stringReference5.Append(stringCache5);

        C.ChartShapeProperties chartShapeProperties4 = new C.ChartShapeProperties();

        A.SolidFill solidFill13 = new A.SolidFill();
        A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent3 };

        solidFill13.Append(schemeColor13);

        A.Outline outline4 = new A.Outline();
        A.NoFill noFill5 = new A.NoFill();

        outline4.Append(noFill5);
        A.EffectList effectList4 = new A.EffectList();

        chartShapeProperties4.Append(solidFill13);
        chartShapeProperties4.Append(outline4);
        chartShapeProperties4.Append(effectList4);

        C.StringCache stringCache6 = new C.StringCache();
        C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)4U };

        C.StringPoint stringPoint12 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue20 = new C.NumericValue();

        stringPoint12.Append(numericValue20);

        C.StringPoint stringPoint13 = new C.StringPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue21 = new C.NumericValue();

        stringPoint13.Append(numericValue21);

        C.StringPoint stringPoint14 = new C.StringPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue22 = new C.NumericValue();

        stringPoint14.Append(numericValue22);

        C.StringPoint stringPoint15 = new C.StringPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue23 = new C.NumericValue();

        stringPoint15.Append(numericValue23);

        stringCache6.Append(pointCount8);
        stringCache6.Append(stringPoint12);
        stringCache6.Append(stringPoint13);
        stringCache6.Append(stringPoint14);
        stringCache6.Append(stringPoint15);

        // stringReference6.Append(stringCache6);

        C.Values values3 = new C.Values();

        C.NumberReference numberReference3 = new C.NumberReference();
        C.Formula formula9 = new C.Formula();

        C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue24 = new C.NumericValue();
        numericValue24.Text = "2";

        numericPoint9.Append(numericValue24);

        C.NumericValue numericValue25 = new C.NumericValue();
        numericValue25.Text = "2";

        numberReference3.Append(formula9);

        values3.Append(numberReference3);

        C.BarSerExtension barSerExtension3 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension3.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");


        C.DataLabels dataLabels1 = new C.DataLabels();
        C.ShowLegendKey showLegendKey1 = new C.ShowLegendKey() { Val = false };
        C.ShowValue showValue1 = new C.ShowValue() { Val = false };
        C.ShowCategoryName showCategoryName1 = new C.ShowCategoryName() { Val = false };
        C.ShowSeriesName showSeriesName1 = new C.ShowSeriesName() { Val = false };
        C.ShowPercent showPercent1 = new C.ShowPercent() { Val = false };
        C.ShowBubbleSize showBubbleSize1 = new C.ShowBubbleSize() { Val = false };

        dataLabels1.Append(showLegendKey1);
        dataLabels1.Append(showValue1);
        dataLabels1.Append(showCategoryName1);
        dataLabels1.Append(showSeriesName1);
        dataLabels1.Append(showPercent1);
        dataLabels1.Append(showBubbleSize1);
        C.GapWidth gapWidth1 = new C.GapWidth() { Val = (UInt16Value)182U };
        C.AxisId axisId1 = new C.AxisId() { Val = (UInt32Value)2020378015U };
        C.AxisId axisId2 = new C.AxisId() { Val = (UInt32Value)2020386175U };

        barChart.Append(barDirection1);
        barChart.Append(barGrouping1);
        barChart.Append(varyColors1);
        barChart.Append(barChartSeriesB);
        barChart.Append(barChartSeriesC);
        barChart.Append(dataLabels1);
        barChart.Append(gapWidth1);
        barChart.Append(axisId1);
        barChart.Append(axisId2);

        C.CategoryAxis categoryAxis1 = new C.CategoryAxis();
        C.AxisId axisId3 = new C.AxisId() { Val = (UInt32Value)2020378015U };

        C.Scaling scaling1 = new C.Scaling();
        C.Orientation orientation1 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        scaling1.Append(orientation1);
        C.Delete delete1 = new C.Delete() { Val = false };
        C.AxisPosition axisPosition1 = new C.AxisPosition() { Val = C.AxisPositionValues.Left };
        C.NumberingFormat numberingFormat1 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
        C.MajorTickMark majorTickMark1 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
        C.MinorTickMark minorTickMark1 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
        C.TickLabelPosition tickLabelPosition1 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

        C.ChartShapeProperties chartShapeProperties5 = new C.ChartShapeProperties();
        A.NoFill noFill6 = new A.NoFill();

        A.Outline outline5 = new A.Outline()
        {
            Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single,
            Alignment = A.PenAlignmentValues.Center
        };

        A.SolidFill solidFill14 = new A.SolidFill();

        A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 15000 };
        A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 85000 };

        schemeColor14.Append(luminanceModulation2);
        schemeColor14.Append(luminanceOffset2);

        solidFill14.Append(schemeColor14);
        A.Round round1 = new A.Round();

        outline5.Append(solidFill14);
        outline5.Append(round1);
        A.EffectList effectList5 = new A.EffectList();

        chartShapeProperties5.Append(noFill6);
        chartShapeProperties5.Append(outline5);
        chartShapeProperties5.Append(effectList5);

        C.TextProperties textProperties2 = new C.TextProperties();
        A.BodyProperties bodyProperties2 = new A.BodyProperties()
        {
            Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true
        };
        A.ListStyle listStyle2 = new A.ListStyle();

        A.Paragraph paragraph2 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties()
        {
            FontSize = 1197, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0
        };

        A.SolidFill solidFill15 = new A.SolidFill();

        A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor15.Append(luminanceModulation3);
        schemeColor15.Append(luminanceOffset3);

        solidFill15.Append(schemeColor15);
        A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties12.Append(solidFill15);
        defaultRunProperties12.Append(latinFont11);
        defaultRunProperties12.Append(eastAsianFont11);
        defaultRunProperties12.Append(complexScriptFont11);

        paragraphProperties2.Append(defaultRunProperties12);
        A.EndParagraphRunProperties endParagraphRunProperties2 =
            new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph2.Append(paragraphProperties2);
        paragraph2.Append(endParagraphRunProperties2);

        textProperties2.Append(bodyProperties2);
        textProperties2.Append(listStyle2);
        textProperties2.Append(paragraph2);
        C.CrossingAxis crossingAxis1 = new C.CrossingAxis() { Val = (UInt32Value)2020386175U };
        C.Crosses crosses1 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        C.AutoLabeled autoLabeled1 = new C.AutoLabeled() { Val = true };
        C.LabelAlignment labelAlignment1 = new C.LabelAlignment() { Val = C.LabelAlignmentValues.Center };
        C.LabelOffset labelOffset1 = new C.LabelOffset() { Val = (UInt16Value)100U };
        C.NoMultiLevelLabels noMultiLevelLabels1 = new C.NoMultiLevelLabels() { Val = false };

        categoryAxis1.Append(axisId3);
        categoryAxis1.Append(scaling1);
        categoryAxis1.Append(delete1);
        categoryAxis1.Append(axisPosition1);
        categoryAxis1.Append(numberingFormat1);
        categoryAxis1.Append(majorTickMark1);
        categoryAxis1.Append(minorTickMark1);
        categoryAxis1.Append(tickLabelPosition1);
        categoryAxis1.Append(chartShapeProperties5);
        categoryAxis1.Append(textProperties2);
        categoryAxis1.Append(crossingAxis1);
        categoryAxis1.Append(crosses1);
        categoryAxis1.Append(autoLabeled1);
        categoryAxis1.Append(labelAlignment1);
        categoryAxis1.Append(labelOffset1);
        categoryAxis1.Append(noMultiLevelLabels1);

        C.ValueAxis valueAxis1 = new C.ValueAxis();
        C.AxisId axisId4 = new C.AxisId() { Val = (UInt32Value)2020386175U };

        C.Scaling scaling2 = new C.Scaling();
        C.Orientation orientation2 = new C.Orientation() { Val = C.OrientationValues.MinMax };

        scaling2.Append(orientation2);
        C.Delete delete2 = new C.Delete() { Val = false };
        C.AxisPosition axisPosition2 = new C.AxisPosition() { Val = C.AxisPositionValues.Bottom };

        C.MajorGridlines majorGridlines1 = new C.MajorGridlines();

        C.ChartShapeProperties chartShapeProperties6 = new C.ChartShapeProperties();

        A.Outline outline6 = new A.Outline()
        {
            Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single,
            Alignment = A.PenAlignmentValues.Center
        };

        A.SolidFill solidFill16 = new A.SolidFill();

        A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 15000 };
        A.LuminanceOffset luminanceOffset4 = new A.LuminanceOffset() { Val = 85000 };

        schemeColor16.Append(luminanceModulation4);
        schemeColor16.Append(luminanceOffset4);

        solidFill16.Append(schemeColor16);
        A.Round round2 = new A.Round();

        outline6.Append(solidFill16);
        outline6.Append(round2);
        A.EffectList effectList6 = new A.EffectList();

        chartShapeProperties6.Append(outline6);
        chartShapeProperties6.Append(effectList6);

        majorGridlines1.Append(chartShapeProperties6);
        C.NumberingFormat numberingFormat2 = new C.NumberingFormat() { FormatCode = "General", SourceLinked = true };
        C.MajorTickMark majorTickMark2 = new C.MajorTickMark() { Val = C.TickMarkValues.None };
        C.MinorTickMark minorTickMark2 = new C.MinorTickMark() { Val = C.TickMarkValues.None };
        C.TickLabelPosition tickLabelPosition2 = new C.TickLabelPosition() { Val = C.TickLabelPositionValues.NextTo };

        C.ChartShapeProperties chartShapeProperties7 = new C.ChartShapeProperties();
        A.NoFill noFill7 = new A.NoFill();

        A.Outline outline7 = new A.Outline();
        A.NoFill noFill8 = new A.NoFill();

        outline7.Append(noFill8);
        A.EffectList effectList7 = new A.EffectList();

        chartShapeProperties7.Append(noFill7);
        chartShapeProperties7.Append(outline7);
        chartShapeProperties7.Append(effectList7);

        C.TextProperties textProperties3 = new C.TextProperties();
        A.BodyProperties bodyProperties3 = new A.BodyProperties()
        {
            Rotation = -60000000, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true
        };
        A.ListStyle listStyle3 = new A.ListStyle();

        A.Paragraph paragraph3 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties()
        {
            FontSize = 1197, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0
        };

        A.SolidFill solidFill17 = new A.SolidFill();

        A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset5 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor17.Append(luminanceModulation5);
        schemeColor17.Append(luminanceOffset5);

        solidFill17.Append(schemeColor17);
        A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties13.Append(solidFill17);
        defaultRunProperties13.Append(latinFont12);
        defaultRunProperties13.Append(eastAsianFont12);
        defaultRunProperties13.Append(complexScriptFont12);

        paragraphProperties3.Append(defaultRunProperties13);
        A.EndParagraphRunProperties endParagraphRunProperties3 =
            new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph3.Append(paragraphProperties3);
        paragraph3.Append(endParagraphRunProperties3);

        textProperties3.Append(bodyProperties3);
        textProperties3.Append(listStyle3);
        textProperties3.Append(paragraph3);
        C.CrossingAxis crossingAxis2 = new C.CrossingAxis() { Val = (UInt32Value)2020378015U };
        C.Crosses crosses2 = new C.Crosses() { Val = C.CrossesValues.AutoZero };
        C.CrossBetween crossBetween1 = new C.CrossBetween() { Val = C.CrossBetweenValues.Between };

        valueAxis1.Append(axisId4);
        valueAxis1.Append(scaling2);
        valueAxis1.Append(delete2);
        valueAxis1.Append(axisPosition2);
        valueAxis1.Append(majorGridlines1);
        valueAxis1.Append(numberingFormat2);
        valueAxis1.Append(majorTickMark2);
        valueAxis1.Append(minorTickMark2);
        valueAxis1.Append(tickLabelPosition2);
        valueAxis1.Append(chartShapeProperties7);
        valueAxis1.Append(textProperties3);
        valueAxis1.Append(crossingAxis2);
        valueAxis1.Append(crosses2);
        valueAxis1.Append(crossBetween1);

        C.ShapeProperties shapeProperties1 = new C.ShapeProperties();
        A.NoFill noFill9 = new A.NoFill();

        A.Outline outline8 = new A.Outline();
        A.NoFill noFill10 = new A.NoFill();

        outline8.Append(noFill10);
        A.EffectList effectList8 = new A.EffectList();

        shapeProperties1.Append(noFill9);
        shapeProperties1.Append(outline8);
        shapeProperties1.Append(effectList8);

        plotArea1.Append(layout1);
        plotArea1.Append(barChart);
        plotArea1.Append(categoryAxis1);
        plotArea1.Append(valueAxis1);
        plotArea1.Append(shapeProperties1);

        C.Legend legend1 = new C.Legend();
        C.LegendPosition legendPosition1 = new C.LegendPosition() { Val = C.LegendPositionValues.Bottom };
        C.Overlay overlay2 = new C.Overlay() { Val = false };

        C.ChartShapeProperties chartShapeProperties8 = new C.ChartShapeProperties();
        A.NoFill noFill11 = new A.NoFill();

        A.Outline outline9 = new A.Outline();
        A.NoFill noFill12 = new A.NoFill();

        outline9.Append(noFill12);
        A.EffectList effectList9 = new A.EffectList();

        chartShapeProperties8.Append(noFill11);
        chartShapeProperties8.Append(outline9);
        chartShapeProperties8.Append(effectList9);

        C.TextProperties textProperties4 = new C.TextProperties();
        A.BodyProperties bodyProperties4 = new A.BodyProperties()
        {
            Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
            Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square,
            Anchor = A.TextAnchoringTypeValues.Center, AnchorCenter = true
        };
        A.ListStyle listStyle4 = new A.ListStyle();

        A.Paragraph paragraph4 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties();

        A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties()
        {
            FontSize = 1197, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None,
            Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Baseline = 0
        };

        A.SolidFill solidFill18 = new A.SolidFill();

        A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
        A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 65000 };
        A.LuminanceOffset luminanceOffset6 = new A.LuminanceOffset() { Val = 35000 };

        schemeColor18.Append(luminanceModulation6);
        schemeColor18.Append(luminanceOffset6);

        solidFill18.Append(schemeColor18);
        A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };
        A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "+mn-ea" };
        A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

        defaultRunProperties14.Append(solidFill18);
        defaultRunProperties14.Append(latinFont13);
        defaultRunProperties14.Append(eastAsianFont13);
        defaultRunProperties14.Append(complexScriptFont13);

        paragraphProperties4.Append(defaultRunProperties14);
        A.EndParagraphRunProperties endParagraphRunProperties4 =
            new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph4.Append(paragraphProperties4);
        paragraph4.Append(endParagraphRunProperties4);

        textProperties4.Append(bodyProperties4);
        textProperties4.Append(listStyle4);
        textProperties4.Append(paragraph4);

        legend1.Append(legendPosition1);
        legend1.Append(overlay2);
        legend1.Append(chartShapeProperties8);
        legend1.Append(textProperties4);
        C.PlotVisibleOnly plotVisibleOnly1 = new C.PlotVisibleOnly() { Val = true };
        C.DisplayBlanksAs displayBlanksAs1 = new C.DisplayBlanksAs() { Val = C.DisplayBlanksAsValues.Gap };

        C.ExtensionList extensionList1 = new C.ExtensionList();

        C.Extension extension1 = new C.Extension() { Uri = "{56B9EC1D-385E-4148-901F-78D8002777C0}" };
        extension1.AddNamespaceDeclaration("c16r3", "http://schemas.microsoft.com/office/drawing/2017/03/chart");

        // OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement(
        //     "<c16r3:dataDisplayOptions16 xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\"><c16r3:dispNaAsBlank val=\"1\" /></c16r3:dataDisplayOptions16>");
        // extension1.Append(openXmlUnknownElement5);

        extensionList1.Append(extension1);
        C.ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new C.ShowDataLabelsOverMaximum() { Val = false };

        cChart.Append(title);
        cChart.Append(autoTitleDeleted1);
        cChart.Append(plotArea1);
        cChart.Append(legend1);
        cChart.Append(plotVisibleOnly1);
        cChart.Append(displayBlanksAs1);
        cChart.Append(extensionList1);
        cChart.Append(showDataLabelsOverMaximum1);

        C.ShapeProperties shapeProperties2 = new C.ShapeProperties();
        A.NoFill noFill13 = new A.NoFill();

        A.Outline outline10 = new A.Outline();
        A.NoFill noFill14 = new A.NoFill();

        outline10.Append(noFill14);
        A.EffectList effectList10 = new A.EffectList();

        shapeProperties2.Append(noFill13);
        shapeProperties2.Append(outline10);
        shapeProperties2.Append(effectList10);

        C.TextProperties textProperties5 = new C.TextProperties();
        A.BodyProperties bodyProperties5 = new A.BodyProperties();
        A.ListStyle listStyle5 = new A.ListStyle();

        A.Paragraph paragraph5 = new A.Paragraph();

        A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties();
        A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties();

        paragraphProperties5.Append(defaultRunProperties15);
        A.EndParagraphRunProperties endParagraphRunProperties5 =
            new A.EndParagraphRunProperties() { Language = "en-US" };

        paragraph5.Append(paragraphProperties5);
        paragraph5.Append(endParagraphRunProperties5);

        textProperties5.Append(bodyProperties5);
        textProperties5.Append(listStyle5);
        textProperties5.Append(paragraph5);

        C.ExternalData externalData1 = new C.ExternalData() { Id = "rId3" };
        C.AutoUpdate autoUpdate1 = new C.AutoUpdate() { Val = false };

        externalData1.Append(autoUpdate1);

        chartSpace.Append(date19041);
        chartSpace.Append(editingLanguage1);
        chartSpace.Append(roundedCorners1);
        chartSpace.Append(alternateContent1);
        chartSpace.Append(cChart);
        chartSpace.Append(shapeProperties2);
        chartSpace.Append(textProperties5);
        chartSpace.Append(externalData1);

        chartPart1.ChartSpace = chartSpace;
    }

    private C.NumericPoint CreateCNumericPoint(int index, int seriesValue)
    {
        var cNumPoint = new C.NumericPoint { Index = (uint)index };
        var cNumValue = new C.NumericValue();
        cNumValue.Text = seriesValue.ToString();
        cNumPoint.Append(cNumValue);

        return cNumPoint;
    }
}