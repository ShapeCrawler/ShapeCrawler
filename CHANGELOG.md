# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## Version 0.4.0 - Unreleased
### Added
- Added setters for `X`, `Y`, `Width` and `Height` properties of non-placeholder shapes;
- Added `ShapeEx.IsGrouped` boolean property to determine whether the shape is grouped.

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
