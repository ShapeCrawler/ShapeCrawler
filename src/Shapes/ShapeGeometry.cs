using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeGeometry(P.ShapeProperties pShapeProperties): IShapeGeometry
{
    /// <summary>
    ///     Corner size on new rounded rectangles, before adjustments are applied.
    /// </summary>
    /// <remarks>
    ///     Rounded rectangles always have a corner size. When they are first created, they have no
    ///     adjustments. The visual appearance of an unadjusted corner is the same as a corner
    ///     with a size of this value.
    /// </remarks>
    private const decimal DefaultCornerSize = 35m;

    /// <summary>
    ///     Mapping of geometries to shape types in outlying cases.
    /// </summary>
    /// <remarks>
    ///     Most geometry types use the same names in both types, or can be programatically
    ///     mapped with substring substitution. These are the outliers which require special
    ///     handling.
    /// </remarks>
    private static readonly Dictionary<Geometry, ShapeTypeValues> GeometryToShapeTypeValuesMap = new()
    {
        { Geometry.RoundedRectangle, ShapeTypeValues.RoundRectangle },
        { Geometry.SingleCornerRoundedRectangle, ShapeTypeValues.Round1Rectangle },
        { Geometry.TopCornersRoundedRectangle, ShapeTypeValues.Round2SameRectangle },
        { Geometry.DiagonalCornersRoundedRectangle, ShapeTypeValues.Round2DiagonalRectangle },
        { Geometry.UTurnArrow, ShapeTypeValues.UTurnArrow },
        { Geometry.LineInverse, ShapeTypeValues.LineInverse },
        { Geometry.RightTriangle, ShapeTypeValues.RightTriangle },
    };

    private static readonly Dictionary<ShapeTypeValues, Geometry> ShapeTypeValuesToGeometryMap
        = GeometryToShapeTypeValuesMap.ToDictionary(x => x.Value, x => x.Key);

    /// <summary>
    ///     Mapping of geometries to the number of adjustments it's expected to have.
    /// </summary>
    /// <remarks>
    ///     Only geometries listed here allow setting adjustments.
    /// </remarks>
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

    public Geometry GeometryType
    {
        get
        {
            var preset = this.APresetGeometry?.Preset;
            if (preset is null)
            {
                if (pShapeProperties.OfType<CustomGeometry>().Any())
                {
                    return Geometry.Custom;
                }

                return Geometry.Rectangle;
            }

            if (!ShapeTypeValuesToGeometryMap.TryGetValue(preset, out var geometryType))
            {
                var presetString = preset.ToString() !;
                var name = presetString.ToLowerInvariant().Replace("rect", "rectangle").Replace("diag", "diagonal");
                return (Geometry)Enum.Parse(typeof(Geometry), name, true);
            }

            return geometryType;
        }

        set
        {
            if (value == Geometry.Custom)
            {
                throw new SCException("Can't set custom geometry");
            }

            var aPresetGeometry = this.APresetGeometry;
            if (aPresetGeometry?.Preset is null && pShapeProperties.OfType<CustomGeometry>().Any())
            {
                throw new SCException("Can't set new geometry on a shape with custom geometry");
            }

            aPresetGeometry ??= pShapeProperties.InsertAt<PresetGeometry>(new(), 0)
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
            aPresetGeometry.RemoveAllChildren<AdjustValueList>();
            aPresetGeometry.AppendChild<AdjustValueList>(new());
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
                (Geometry.TopCornersRoundedRectangle, 0) => DefaultCornerSize,
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
                Geometry.TopCornersRoundedRectangle => [value, 0],
                _ => throw new SCException($"{geometryType} does not support {nameof(this.CornerSize)}")
            };
        }
    }

    public decimal[] Adjustments
    {
        get => this.ExtractAdjustmentsFromShapeGuide();
        set
        {
            if (GeometryToNumberOfAdjustmentsMap.TryGetValue(this.GeometryType, out var numAdjustments))
            {
                if (value.Length > numAdjustments)
                {
                    throw new SCException($"{this.GeometryType} only supports {numAdjustments} adjustments");
                }

                if (value.Length < numAdjustments && this.ExtractAdjustmentsFromShapeGuide().Length < numAdjustments)
                {
                    // If user is not setting sufficient quantity of adjustments, AND there are
                    // not already enough adjustments in place, we need to resize up to the
                    // total expected number of adjustments.
                    Array.Resize(ref value, numAdjustments);
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
                throw new SCException($"{this.GeometryType} does not support adjustments");
            }
        }
    }

    private PresetGeometry? APresetGeometry => pShapeProperties.GetFirstChild<PresetGeometry>();

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

    private static decimal ExtractAdjustmentFromShapeGuide(ShapeGuide sg)
    {
        var formula = sg.Formula?.Value ?? throw new SCException("Malformed geometry. Shape guide has no formula.");

        var pattern = "^val (?<value>-?[0-9]+)$";

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
            .Descendants<ShapeGuide>()
            .Where(x => x.Name?.Value?.StartsWith("adj", StringComparison.InvariantCulture) ?? false)
            .OrderBy(x => x.Name?.Value ?? string.Empty)
            .Select(ExtractAdjustmentFromShapeGuide)
            .ToArray()
            ?? throw new SCException("Malformed geometry.");
    }

    private void InjectSingleAdjustmentToShapeGuide(decimal[] values)
    {
        if (values.Length != 1)
        {
            throw new SCException("This geometry supports only a single adjustment value.");
        }

        this.Inject("adj", values[0]);
    }

    private void InjectMultipleAdjustmentsIntoShapeGuide(decimal[] values)
    {
        for (var i = 0; i < values.Length; i++)
        {
            this.Inject($"adj{i + 1}", values[i]);
        }
    }

    private void Inject(string name, decimal value)
    {
        var avList = this.APresetGeometry?.AdjustValueList
            ?? throw new SCException("Malformed geometry. Missing AdjustValueList.");

        var sgs = avList
            .Descendants<ShapeGuide>()
            .Where(x => x.Name == name);

        if (sgs.Count() > 1)
        {
            throw new SCException($"Malformed geometry. Has multiple {name} shape guides.");
        }

        // Will add a shape guide if there isn't already one
        var sg = sgs.SingleOrDefault()
            ?? avList.AppendChild(new ShapeGuide() { Name = name })
            ?? throw new SCException("Failed attempting to add a shape guide to AdjustValueList");

        sg.Formula = new StringValue($"val {(int)(value * 500m)}");
    }
}