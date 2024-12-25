using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class ShapeGeometry : IShapeGeometry
{
    private const string ExceptionMessageMissingAdjustValueList = "Malformed rounded rectangle. Missing AdjustValueList.";

    private static readonly Dictionary<Geometry, ShapeTypeValues> GeometryToShapeTypeValuesMap = new()
    {
        { Geometry.RoundedRectangle, A.ShapeTypeValues.RoundRectangle },
        { Geometry.SingleCornerRoundedRectangle, A.ShapeTypeValues.Round1Rectangle },
        { Geometry.TopCornersRoundedRectangle, A.ShapeTypeValues.Round2SameRectangle },
        { Geometry.DiagonalCornersRoundedRectangle, A.ShapeTypeValues.Round2DiagonalRectangle },
        { Geometry.UTurnArrow, A.ShapeTypeValues.UTurnArrow },
        { Geometry.LineInverse, A.ShapeTypeValues.LineInverse },
        { Geometry.RightTriangle, A.ShapeTypeValues.RightTriangle },
    };
    
    private static readonly Dictionary<ShapeTypeValues, Geometry> ShapeTypeValuesToGeometryMap 
        = GeometryToShapeTypeValuesMap.ToDictionary(x => x.Value, x => x.Key);

    private static readonly Dictionary<Geometry, int> GeometryToNumberOfAdjustmentsMap = new()
    {
        { Geometry.RoundedRectangle, 1 },
        { Geometry.SingleCornerRoundedRectangle, 1 },
        { Geometry.TopCornersRoundedRectangle, 2 },
        { Geometry.DiagonalCornersRoundedRectangle, 2 },
        { Geometry.Snip1Rectangle, 1 },
        { Geometry.Snip2DiagonalRectangle, 2 },
        { Geometry.Snip2SameRectangle, 2 },
        { Geometry.SnipRoundRectangle, 2 },
    };

    private readonly P.ShapeProperties pShapeProperties;

    internal ShapeGeometry(P.ShapeProperties pShapeProperties)
    {
        this.pShapeProperties = pShapeProperties;
    }

    public Geometry GeometryType 
    { 
        get
        {
            var preset = this.APresetGeometry?.Preset;
            if (preset is null)
            {
                if (this.pShapeProperties.OfType<A.CustomGeometry>().Any())
                {
                    return Geometry.Custom;
                }
                else
                {
                    return Geometry.Rectangle;
                }
            }
            else
            {
                if (!ShapeTypeValuesToGeometryMap.TryGetValue(preset, out Geometry geometryType))
                {
                    var presetString = preset!.ToString() !;
                    var name = presetString.ToLowerInvariant().Replace("rect", "rectangle").Replace("diag", "diagonal");
                    return (Geometry)Enum.Parse(typeof(Geometry), name, true);
                }

                return geometryType;
            }            
        }
        
        set
        {
            if (value == Geometry.Custom)
            {
                throw new SCException("Can't set custom geometry");
            }

            var aPresetGeometry = this.APresetGeometry;
            if (aPresetGeometry?.Preset is null && this.pShapeProperties.OfType<A.CustomGeometry>().Any())
            {
                throw new SCException("Can't set new geometry on a shape with custom geometry");
            }

            aPresetGeometry ??= this.pShapeProperties.InsertAt<A.PresetGeometry>(new(), 0)
                ?? throw new SCException("Unable to add new preset geometry");

            if (!GeometryToShapeTypeValuesMap.TryGetValue(value, out var newPreset))
            {
                var name = value.ToString().Replace("Rectangle", "Rect").Replace("Diagonal", "Diag");
                var camelName = ToCamelCaseInvariant(name);        
                newPreset = new ShapeTypeValues(camelName);
            }

            if (!(newPreset as IEnumValue).IsValid)
            {
                throw new SCException($"Invalid preset value {newPreset}");
            }

            aPresetGeometry.Preset = newPreset;

            // Presets have different expectations of an adjusted value lists, so changing the
            // preset means we must remove any existing adjustments, and create a new empty one
            aPresetGeometry.RemoveAllChildren<A.AdjustValueList>();
            aPresetGeometry.AppendChild<A.AdjustValueList>(new());
        }
    }

    public decimal CornerSize
    {
        get
        {
            var adjustments = this.Adjustments;
            return (this.GeometryType, adjustments.Length) switch
            {
                (Geometry.RoundedRectangle, 0) or 
                (Geometry.TopCornersRoundedRectangle, 0) => 35,
                (Geometry.RoundedRectangle, _) or 
                (Geometry.TopCornersRoundedRectangle, _) => adjustments[0],
                _ => 0
            };
        }
        
        set
        {
            var geometryType = this.GeometryType;
            this.Adjustments = geometryType switch
            {
                Geometry.RoundedRectangle => [value],
                Geometry.TopCornersRoundedRectangle => [value,0],
                _ => throw new SCException($"{geometryType} does not support {nameof(CornerSize)}")
            };
        }
    }

    /// <summary>
    ///     Gets or sets the geometry adjustments. Work in progress!! 
    /// </summary>
    internal decimal[] Adjustments
    {
        get => ExtractAdjustmentsFromShapeGuide();
        set {
            if (GeometryToNumberOfAdjustmentsMap.TryGetValue(GeometryType, out var numAdjustments))
            {
                if (value.Length > numAdjustments)
                {
                    throw new SCException($"{GeometryType} only supports {numAdjustments} adjustments");
                }

                if (numAdjustments == 1)
                {
                    this.InjectSingleAdjustmentToShapeGuide(value);
                }
                else
                {
                    this.InjectMultipleAdjustmentsIntoShapeGuide(value);
                }
            }
            else
            {
                throw new SCException($"{GeometryType} does not support adjustments");
            }
        }
    }

    private A.PresetGeometry? APresetGeometry => this.pShapeProperties.GetFirstChild<A.PresetGeometry>();

    internal void UpdateGeometry(Geometry type)
    {
        this.GeometryType = type;
    }

    private static string ToCamelCaseInvariant(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        if (value.Length == 1)
        {
            return value.ToLowerInvariant();
        }

        return char.ToLowerInvariant(value[0]) + value[1..];
    }

    private static decimal ExtractAdjustmentFromShapeGuide(A.ShapeGuide sg)
    {
        var formula = sg.Formula?.Value ?? throw new SCException("Malformed geometry. Shape guide has no formula.");

        var pattern = "^val (?<value>[0-9]+)$";

#if NETSTANDARD2_0
        var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
        var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

        var match = regex.Match(formula);
        if (!match.Success)
        {
            throw new SCException("Malformed geometry. Formula has no value.");
        }

        var value = int.Parse(match.Groups["value"].Value);

        return value / 500m;
    }

    private decimal[] ExtractAdjustmentsFromShapeGuide()
    {
        return this.APresetGeometry?
            .AdjustValueList?
            .Descendants<A.ShapeGuide>()
            .Where(x => x.Name?.Value?.StartsWith("adj") ?? false)
            .OrderBy(x => x.Name?.Value ?? string.Empty)
            .Select(ExtractAdjustmentFromShapeGuide)
            .ToArray()
            ?? throw new SCException("Malformed geometry.");
    }

    private void InjectSingleAdjustmentToShapeGuide(decimal[] values)
    {
        if (values.Length != 1)
        {
            throw new SCException("This geometry supports a single adjustment value.");
        }

        Inject("adj", values[0]);
    }

    private void InjectMultipleAdjustmentsIntoShapeGuide(decimal[] values)
    {
        for (int i = 0; i < values.Length; i++)
        {
            Inject($"adj{i + 1}", values[i]);
        }
    }

    private void Inject(string name, decimal value)
    {
        var avList = this.APresetGeometry?.AdjustValueList 
            ?? throw new SCException(ExceptionMessageMissingAdjustValueList);

        var sgs = avList
            .Descendants<A.ShapeGuide>()
            .Where(x => x.Name == name);

        if (sgs.Count() > 1)
        {
            throw new SCException($"Malformed geometry. Has multiple {name} shape guides.");
        }

        // Will add a shape guide if there isn't already one
        var sg = sgs.SingleOrDefault()
            ?? avList.AppendChild(new A.ShapeGuide() { Name = name }) 
            ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");

        sg.Formula = new StringValue($"val {(int)(value * 500m)}");        
    }
}