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

    public P.GraphicFrame Create(TypedOpenXmlPart typedOpenXmlPart)
    {
        var id = (UInt32Value)6U;
        var name = "Chart 5";
        
        var chartPart = typedOpenXmlPart.AddNewPart<ChartPart>("rId2");
        this.GenerateChartPartContent(chartPart);

        // Create Excel
        var embeddedPackagePart = chartPart.AddNewPart<EmbeddedPackagePart>(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId3");
        using var spreadsheetDocument = SpreadsheetDocument.Create(embeddedPackagePart.GetStream(FileMode.Create),
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
        
        var row = new X.Row { RowIndex = 1 };
        var cell = new X.Cell
        {
            CellReference = "B1",
            DataType = X.CellValues.String,
            CellValue = new X.CellValue("Series 1")
        };
        row.Append(cell);
        sheetData.Append(row);
        spreadsheetDocument.Save();
        spreadsheetDocument.Dispose();

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
        chartReference.AddNamespaceDeclaration("r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        graphicData.Append(chartReference);

        graphic.Append(graphicData);

        graphicFrame.Append(nonVisualGraphicFrameProperties);
        graphicFrame.Append(transform);
        graphicFrame.Append(graphic);

        return graphicFrame;
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
        alternateContentChoice1.AddNamespaceDeclaration("c14",
            "http://schemas.microsoft.com/office/drawing/2007/8/2/chart");
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

        C.BarChart barChart1 = new C.BarChart();
        C.BarDirection barDirection1 = new C.BarDirection() { Val = C.BarDirectionValues.Bar };
        C.BarGrouping barGrouping1 = new C.BarGrouping() { Val = C.BarGroupingValues.Clustered };
        C.VaryColors varyColors1 = new C.VaryColors() { Val = false };

        C.BarChartSeries barChartSeries1 = new C.BarChartSeries();
        C.Index index1 = new C.Index() { Val = (UInt32Value)0U };
        C.Order order1 = new C.Order() { Val = (UInt32Value)0U };

        var series1TitleText = new C.SeriesText();
        var series1TitleRef = new C.StringReference();
        var series1TitleFormula = new C.Formula();
        series1TitleFormula.Text = "Sheet1!$B$1";
        series1TitleRef.Append(series1TitleFormula);
        series1TitleText.Append(series1TitleRef);

        var stringCache1 = new C.StringCache();
        var pointCount1 = new C.PointCount { Val = (UInt32Value)1U };
        var stringPoint1 = new C.StringPoint { Index = (UInt32Value)0U };
        var numericValue1 = new C.NumericValue();
        numericValue1.Text = "Series 1";
        stringPoint1.Append(numericValue1);
        stringCache1.Append(pointCount1);
        stringCache1.Append(stringPoint1);
        series1TitleRef.Append(stringCache1);
        

        C.ChartShapeProperties chartShapeProperties2 = new C.ChartShapeProperties();

        A.SolidFill solidFill11 = new A.SolidFill();
        A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

        solidFill11.Append(schemeColor11);

        A.Outline outline2 = new A.Outline();
        A.NoFill noFill3 = new A.NoFill();

        outline2.Append(noFill3);
        A.EffectList effectList2 = new A.EffectList();

        chartShapeProperties2.Append(solidFill11);
        chartShapeProperties2.Append(outline2);
        chartShapeProperties2.Append(effectList2);
        C.InvertIfNegative invertIfNegative1 = new C.InvertIfNegative() { Val = false };

        C.CategoryAxisData categoryAxisData1 = new C.CategoryAxisData();

        C.StringReference stringReference2 = new C.StringReference();
        C.Formula formula2 = new C.Formula();
        formula2.Text = "Sheet1!$A$2:$A$5";

        C.StringCache stringCache2 = new C.StringCache();
        C.PointCount pointCount2 = new C.PointCount() { Val = (UInt32Value)4U };

        C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue2 = new C.NumericValue();
        numericValue2.Text = "Category 1";

        stringPoint2.Append(numericValue2);

        C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue3 = new C.NumericValue();
        numericValue3.Text = "Category 2";

        stringPoint3.Append(numericValue3);

        C.StringPoint stringPoint4 = new C.StringPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue4 = new C.NumericValue();
        numericValue4.Text = "Category 3";

        stringPoint4.Append(numericValue4);

        C.StringPoint stringPoint5 = new C.StringPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue5 = new C.NumericValue();
        numericValue5.Text = "Category 4";

        stringPoint5.Append(numericValue5);

        stringCache2.Append(pointCount2);
        stringCache2.Append(stringPoint2);
        stringCache2.Append(stringPoint3);
        stringCache2.Append(stringPoint4);
        stringCache2.Append(stringPoint5);

        stringReference2.Append(formula2);
        stringReference2.Append(stringCache2);

        categoryAxisData1.Append(stringReference2);

        C.Values values1 = new C.Values();

        C.NumberReference numberReference1 = new C.NumberReference();
        C.Formula formula3 = new C.Formula();
        formula3.Text = "Sheet1!$B$2:$B$5";

        C.NumberingCache numberingCache1 = new C.NumberingCache();
        C.FormatCode formatCode1 = new C.FormatCode();
        formatCode1.Text = "General";
        C.PointCount pointCount3 = new C.PointCount() { Val = (UInt32Value)4U };

        C.NumericPoint numericPoint1 = new C.NumericPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue6 = new C.NumericValue();
        numericValue6.Text = "4.3";

        numericPoint1.Append(numericValue6);

        C.NumericPoint numericPoint2 = new C.NumericPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue7 = new C.NumericValue();
        numericValue7.Text = "2.5";

        numericPoint2.Append(numericValue7);

        C.NumericPoint numericPoint3 = new C.NumericPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue8 = new C.NumericValue();
        numericValue8.Text = "3.5";

        numericPoint3.Append(numericValue8);

        C.NumericPoint numericPoint4 = new C.NumericPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue9 = new C.NumericValue();
        numericValue9.Text = "4.5";

        numericPoint4.Append(numericValue9);

        numberingCache1.Append(formatCode1);
        numberingCache1.Append(pointCount3);
        numberingCache1.Append(numericPoint1);
        numberingCache1.Append(numericPoint2);
        numberingCache1.Append(numericPoint3);
        numberingCache1.Append(numericPoint4);

        numberReference1.Append(formula3);
        numberReference1.Append(numberingCache1);

        values1.Append(numberReference1);

        C.BarSerExtensionList barSerExtensionList1 = new C.BarSerExtensionList();

        C.BarSerExtension barSerExtension1 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension1.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        barSerExtensionList1.Append(barSerExtension1);

        barChartSeries1.Append(index1);
        barChartSeries1.Append(order1);
        barChartSeries1.Append(series1TitleText);
        barChartSeries1.Append(chartShapeProperties2);
        barChartSeries1.Append(invertIfNegative1);
        barChartSeries1.Append(categoryAxisData1);
        barChartSeries1.Append(values1);
        barChartSeries1.Append(barSerExtensionList1);

        C.BarChartSeries barChartSeries2 = new C.BarChartSeries();
        C.Index index2 = new C.Index() { Val = (UInt32Value)1U };
        C.Order order2 = new C.Order() { Val = (UInt32Value)1U };

        C.SeriesText seriesText2 = new C.SeriesText();

        C.StringReference stringReference3 = new C.StringReference();
        C.Formula formula4 = new C.Formula();
        formula4.Text = "Sheet1!$C$1";

        C.StringCache stringCache3 = new C.StringCache();
        C.PointCount pointCount4 = new C.PointCount() { Val = (UInt32Value)1U };

        C.StringPoint stringPoint6 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue10 = new C.NumericValue();
        numericValue10.Text = "Series 2";

        stringPoint6.Append(numericValue10);

        stringCache3.Append(pointCount4);
        stringCache3.Append(stringPoint6);

        stringReference3.Append(formula4);
        stringReference3.Append(stringCache3);

        seriesText2.Append(stringReference3);

        C.ChartShapeProperties chartShapeProperties3 = new C.ChartShapeProperties();

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

        C.CategoryAxisData categoryAxisData2 = new C.CategoryAxisData();

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

        categoryAxisData2.Append(categoryStringReference4);

        C.Values values2 = new C.Values();

        C.NumberReference numberReference2 = new C.NumberReference();
        C.Formula formula6 = new C.Formula();
        formula6.Text = "Sheet1!$C$2:$C$5";

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

        numberReference2.Append(formula6);
        numberReference2.Append(numberingCache2);

        values2.Append(numberReference2);

        C.BarSerExtensionList barSerExtensionList2 = new C.BarSerExtensionList();

        C.BarSerExtension barSerExtension2 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension2.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        barSerExtensionList2.Append(barSerExtension2);

        barChartSeries2.Append(index2);
        barChartSeries2.Append(order2);
        barChartSeries2.Append(seriesText2);
        barChartSeries2.Append(chartShapeProperties3);
        barChartSeries2.Append(invertIfNegative2);
        barChartSeries2.Append(categoryAxisData2);
        barChartSeries2.Append(values2);
        barChartSeries2.Append(barSerExtensionList2);

        C.BarChartSeries barChartSeries3 = new C.BarChartSeries();
        C.Index index3 = new C.Index() { Val = (UInt32Value)2U };
        C.Order order3 = new C.Order() { Val = (UInt32Value)2U };

        C.SeriesText seriesText3 = new C.SeriesText();

        C.StringReference stringReference5 = new C.StringReference();
        C.Formula formula7 = new C.Formula();
        formula7.Text = "Sheet1!$D$1";

        C.StringCache stringCache5 = new C.StringCache();
        C.PointCount pointCount7 = new C.PointCount() { Val = (UInt32Value)1U };

        C.StringPoint stringPoint11 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue19 = new C.NumericValue();
        numericValue19.Text = "Series 3";

        stringPoint11.Append(numericValue19);

        stringCache5.Append(pointCount7);
        stringCache5.Append(stringPoint11);

        stringReference5.Append(formula7);
        stringReference5.Append(stringCache5);

        seriesText3.Append(stringReference5);

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
        C.InvertIfNegative invertIfNegative3 = new C.InvertIfNegative() { Val = false };

        C.CategoryAxisData categoryAxisData3 = new C.CategoryAxisData();

        C.StringReference stringReference6 = new C.StringReference();
        C.Formula formula8 = new C.Formula();
        formula8.Text = "Sheet1!$A$2:$A$5";

        C.StringCache stringCache6 = new C.StringCache();
        C.PointCount pointCount8 = new C.PointCount() { Val = (UInt32Value)4U };

        C.StringPoint stringPoint12 = new C.StringPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue20 = new C.NumericValue();
        numericValue20.Text = "Category 1";

        stringPoint12.Append(numericValue20);

        C.StringPoint stringPoint13 = new C.StringPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue21 = new C.NumericValue();
        numericValue21.Text = "Category 2";

        stringPoint13.Append(numericValue21);

        C.StringPoint stringPoint14 = new C.StringPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue22 = new C.NumericValue();
        numericValue22.Text = "Category 3";

        stringPoint14.Append(numericValue22);

        C.StringPoint stringPoint15 = new C.StringPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue23 = new C.NumericValue();
        numericValue23.Text = "Category 4";

        stringPoint15.Append(numericValue23);

        stringCache6.Append(pointCount8);
        stringCache6.Append(stringPoint12);
        stringCache6.Append(stringPoint13);
        stringCache6.Append(stringPoint14);
        stringCache6.Append(stringPoint15);

        stringReference6.Append(formula8);
        stringReference6.Append(stringCache6);

        categoryAxisData3.Append(stringReference6);

        C.Values values3 = new C.Values();

        C.NumberReference numberReference3 = new C.NumberReference();
        C.Formula formula9 = new C.Formula();
        formula9.Text = "Sheet1!$D$2:$D$5";

        C.NumberingCache numberingCache3 = new C.NumberingCache();
        C.FormatCode formatCode3 = new C.FormatCode();
        formatCode3.Text = "General";
        C.PointCount pointCount9 = new C.PointCount() { Val = (UInt32Value)4U };

        C.NumericPoint numericPoint9 = new C.NumericPoint() { Index = (UInt32Value)0U };
        C.NumericValue numericValue24 = new C.NumericValue();
        numericValue24.Text = "2";

        numericPoint9.Append(numericValue24);

        C.NumericPoint numericPoint10 = new C.NumericPoint() { Index = (UInt32Value)1U };
        C.NumericValue numericValue25 = new C.NumericValue();
        numericValue25.Text = "2";

        numericPoint10.Append(numericValue25);

        C.NumericPoint numericPoint11 = new C.NumericPoint() { Index = (UInt32Value)2U };
        C.NumericValue numericValue26 = new C.NumericValue();
        numericValue26.Text = "3";

        numericPoint11.Append(numericValue26);

        C.NumericPoint numericPoint12 = new C.NumericPoint() { Index = (UInt32Value)3U };
        C.NumericValue numericValue27 = new C.NumericValue();
        numericValue27.Text = "5";

        numericPoint12.Append(numericValue27);

        numberingCache3.Append(formatCode3);
        numberingCache3.Append(pointCount9);
        numberingCache3.Append(numericPoint9);
        numberingCache3.Append(numericPoint10);
        numberingCache3.Append(numericPoint11);
        numberingCache3.Append(numericPoint12);

        numberReference3.Append(formula9);
        numberReference3.Append(numberingCache3);

        values3.Append(numberReference3);

        C.BarSerExtensionList barSerExtensionList3 = new C.BarSerExtensionList();

        C.BarSerExtension barSerExtension3 = new C.BarSerExtension() { Uri = "{C3380CC4-5D6E-409C-BE32-E72D297353CC}" };
        barSerExtension3.AddNamespaceDeclaration("c16", "http://schemas.microsoft.com/office/drawing/2014/chart");

        barSerExtensionList3.Append(barSerExtension3);

        barChartSeries3.Append(index3);
        barChartSeries3.Append(order3);
        barChartSeries3.Append(seriesText3);
        barChartSeries3.Append(chartShapeProperties4);
        barChartSeries3.Append(invertIfNegative3);
        barChartSeries3.Append(categoryAxisData3);
        barChartSeries3.Append(values3);
        barChartSeries3.Append(barSerExtensionList3);

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

        barChart1.Append(barDirection1);
        barChart1.Append(barGrouping1);
        barChart1.Append(varyColors1);
        barChart1.Append(barChartSeries1);
        barChart1.Append(barChartSeries2);
        barChart1.Append(barChartSeries3);
        barChart1.Append(dataLabels1);
        barChart1.Append(gapWidth1);
        barChart1.Append(axisId1);
        barChart1.Append(axisId2);

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
        plotArea1.Append(barChart1);
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
}