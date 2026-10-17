# Straight line fixture

`straight-lines.pptx` is an Open XML fixture built from the project's blank presentation template, independently of the `ShapeCrawler.AddLine` API. It has two slides with identical coordinates: slide 1 uses the `line` preset and slide 2 uses `straightConnector1`.

Each slide contains a rectangle, seven straight arrow connectors, and a plain horizontal connector. The lines cover horizontal, vertical, and diagonal orientations, including reversed directions and all four combinations of horizontal/vertical flips. The horizontal arrow and plain line have identical endpoints to verify that arrowheads do not affect coordinates.

Coordinates are specified in PowerPoint points and stored as EMUs (12,700 EMUs per point). Fractional point values exercise the conversion without integer rounding. Exact expected endpoints are listed in `LineShapeTests.Straight_connectors_from_pptx_have_line_geometry_and_endpoints_in_points`.

This fixture intentionally excludes rotation, groups, curved/elbow connectors, and connector attachments.
