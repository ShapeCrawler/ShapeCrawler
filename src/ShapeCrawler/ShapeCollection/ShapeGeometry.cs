using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class ShapeGeometry : IShapeGeometry
{
    private readonly OpenXmlElement sdkPShapeTreeElement;

    internal ShapeGeometry(OpenXmlElement sdkPShapeTreeElement)
    {
        this.sdkPShapeTreeElement = sdkPShapeTreeElement;
    }

    public Geometry GeometryType 
    { 
        get
        {
            var aPresetGeometry = this.APresetGeometry();
            var preset = aPresetGeometry?.Preset;
            if (preset is null)
            {
                if (this.PShapeProperties().OfType<A.CustomGeometry>().Any())
                {
                    return Geometry.Custom;
                }
            }
            else
            {                
                // TODO: Reconsider these two clauses. I think they will be picked up fine by the enum tryparse below.
                if(preset.Value == A.ShapeTypeValues.RoundRectangle)
                {
                    return Geometry.RoundRectangle;
                }

                if(preset.Value == A.ShapeTypeValues.Round2SameRectangle)
                {
                    return Geometry.Round2SameRectangle;
                }

                var name = preset.ToString();
                if (name == "rect")
                {
                    return Geometry.Rectangle;
                }

                Enum.TryParse(name, true, out Geometry geometryType);
                return geometryType;    
            }
            
            return Geometry.Rectangle;
        }
        
        set => throw new System.NotImplementedException(); 
    }

    public decimal? CornerSize
    {
        get
        {
            var shapeType = this.APresetGeometry()?.Preset?.Value;

            if (shapeType == A.ShapeTypeValues.RoundRectangle)
            {
                return ExtractCornerSizeFromRoundRectangle(this.APresetGeometry() !);
            }

            if (shapeType == A.ShapeTypeValues.Round2SameRectangle)
            {
                return ExtractCornerSizeFromRound2SameRectangle(this.APresetGeometry() !);
            }

            return null;
        }
        
        set
        {
            if (value is null)
            {
                throw new SCException("Not allowed to set null size. Try 0 to straighten the corner.");
            }

            var shapeType = this.APresetGeometry()?.Preset?.Value;

            if (shapeType == A.ShapeTypeValues.RoundRectangle)
            {
                InjectCornerSizeIntoRoundRectangle(this.APresetGeometry() !, value.Value);
            }

            if (shapeType == A.ShapeTypeValues.Round2SameRectangle)
            {
                InjectCornerSizeIntoRound2SameRectangle(this.APresetGeometry() !, value.Value);
            }
        }
    }

    private static decimal? ExtractCornerSizeFromRoundRectangle(A.PresetGeometry aPresetGeometry)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.RoundRectangle)
        {
            return null;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj");
        if (!sgs.Any())
        {
            // Has no shape guide. That means we're using the DEFAULT value, which is 0.35
            return 0.35m;
        }

        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        return ExtractCornerSizeFromShapeGuide(sgs.Single());
    }

    private static void InjectCornerSizeIntoRoundRectangle(A.PresetGeometry aPresetGeometry, decimal value)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.RoundRectangle)
        {
            return;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj");
        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        // Will add a shape guide if there isn't already one
        var sg = sgs.SingleOrDefault()
            ?? avList.AppendChild(new A.ShapeGuide() { Name = "adj" }) 
            ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");

        sg.Formula = new StringValue($"val {(int)(value * 50000m)}");        
    }

    private static decimal? ExtractCornerSizeFromRound2SameRectangle(A.PresetGeometry aPresetGeometry)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.Round2SameRectangle)
        {
            return null;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>();
        var count = sgs.Count();
        if (count == 0)
        {
            // Has no shape guide. That means we're using the DEFAULT value, which is 0.35
            return 0.35m;
        }

        if (count != 2)
        {
            throw new SCException($"Malformed rounded rectangle. Expected 2 shape guides, found {count}. Please file a GitHub issue.");
        }

        var sg = sgs.SingleOrDefault(x => x.Name == "adj1") ?? throw new SCException($"Malformed rounded rectangle. No shape guide named `adj1`. Please file a GitHub issue.");

        return ExtractCornerSizeFromShapeGuide(sg);
    }

    private static void InjectCornerSizeIntoRound2SameRectangle(A.PresetGeometry aPresetGeometry, decimal value)
    {
        if (aPresetGeometry.Preset?.Value != A.ShapeTypeValues.Round2SameRectangle)
        {
            return;
        }

        var avList = aPresetGeometry.AdjustValueList ?? throw new SCException("Malformed rounded rectangle. Missing AdjustValueList. Please file a GitHub issue.");
        var sgs = avList.Descendants<A.ShapeGuide>().Where(x => x.Name == "adj1");
        if (sgs.Count() > 1)
        {
            throw new SCException("Malformed rounded rectangle. Has multiple shape guides. Please file a GitHub issue.");
        }

        var sg = sgs.SingleOrDefault();
        if (sg is null)
        {
            // Has no adj1 shape guide. We need to add an adj1/adj2 pair
            sg = avList.AppendChild(new A.ShapeGuide() { Name = "adj1" }) ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");
            if (avList.AppendChild(new A.ShapeGuide() { Name = "adj2", Formula = "val 0" }) is null)
            {
                throw new SCException("Failed attempting to add a shape guide to AdjustValueList");
            }
        }
    
        sg.Formula = new StringValue($"val {(int)(value * 50000m)}");        
    }

    private static decimal ExtractCornerSizeFromShapeGuide(A.ShapeGuide sg)
    {
        var formula = sg.Formula?.Value ?? throw new SCException("Malformed rounded rectangle. Shape guide has no formula. Please file a GitHub issue.");

        var pattern = "^val (?<value>[0-9]+)$";

#if NETSTANDARD2_0
        var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
        var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

        var match = regex.Match(formula);
        if (!match.Success)
        {
            throw new SCException("Malformed rounded rectangle. Formula has no value. Please file a GitHub issue.");
        }

        var value = int.Parse(match.Groups["value"].Value);

        return value / 50000m;
    }

    private A.PresetGeometry? APresetGeometry()
    {
        return this.PShapeProperties().GetFirstChild<A.PresetGeometry>();
    }

    private P.ShapeProperties PShapeProperties()
    {
        return this.sdkPShapeTreeElement.Descendants<P.ShapeProperties>().First();
    }
}