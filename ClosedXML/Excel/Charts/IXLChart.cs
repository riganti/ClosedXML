#nullable disable

using System;
using System.Drawing;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Excel
{
    public enum XLChartType {
        Area,
        Area3D,
        AreaStacked,
        AreaStacked100Percent,
        AreaStacked100Percent3D,
        AreaStacked3D,
        BarClustered,
        BarClustered3D,
        BarStacked,
        BarStacked100Percent,
        BarStacked100Percent3D,
        BarStacked3D,
        Bubble,
        Bubble3D,
        Column3D,
        ColumnClustered,
        ColumnClustered3D,
        ColumnStacked,
        ColumnStacked100Percent,
        ColumnStacked100Percent3D,
        ColumnStacked3D,
        Cone,
        ConeClustered,
        ConeHorizontalClustered,
        ConeHorizontalStacked,
        ConeHorizontalStacked100Percent,
        ConeStacked,
        ConeStacked100Percent,
        Cylinder,
        CylinderClustered,
        CylinderHorizontalClustered,
        CylinderHorizontalStacked,
        CylinderHorizontalStacked100Percent,
        CylinderStacked,
        CylinderStacked100Percent,
        Doughnut,
        DoughnutExploded,
        Line,
        Line3D,
        LineStacked,
        LineStacked100Percent,
        LineWithMarkers,
        LineWithMarkersStacked,
        LineWithMarkersStacked100Percent,
        Pie,
        Pie3D,
        PieExploded,
        PieExploded3D,
        PieToBar,
        PieToPie,
        Pyramid,
        PyramidClustered,
        PyramidHorizontalClustered,
        PyramidHorizontalStacked,
        PyramidHorizontalStacked100Percent,
        PyramidStacked,
        PyramidStacked100Percent,
        Radar,
        RadarFilled,
        RadarWithMarkers,
        StockHighLowClose,
        StockOpenHighLowClose,
        StockVolumeHighLowClose,
        StockVolumeOpenHighLowClose,
        Surface,
        SurfaceContour,
        SurfaceContourWireframe,
        SurfaceWireframe,
        XYScatterMarkers,
        XYScatterSmoothLinesNoMarkers,
        XYScatterSmoothLinesWithMarkers,
        XYScatterStraightLinesNoMarkers,
        XYScatterStraightLinesWithMarkers
    }
    public interface IXLChart: IXLDrawing<IXLChart>
    {
        IXLCell BottomRightCell { get; }

        Int32 Id { get; }

        /// <summary>
        /// Current width of the chart in pixels.
        /// </summary>
        Int32 Width { get; set; }

        /// <summary>
        /// Current height of the chart in pixels.
        /// </summary>
        Int32 Height { get; set; }

        Int32 Left { get; set; }

        XLPicturePlacement Placement { get; set; }

        Int32 Top { get; set; }

        IXLCell TopLeftCell { get; }

        Point GetOffset(XLMarkerPosition position);

        IXLChart MoveTo(Int32 left, Int32 top);

        IXLChart MoveTo(IXLCell cell);

        IXLChart MoveTo(IXLCell cell, Int32 xOffset, Int32 yOffset);

        IXLChart MoveTo(IXLCell cell, Point offset);

        IXLChart MoveTo(IXLCell fromCell, IXLCell toCell);

        IXLChart MoveTo(IXLCell fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLCell toCell, Int32 toCellXOffset, Int32 toCellYOffset);

        IXLChart MoveTo(IXLCell fromCell, Point fromOffset, IXLCell toCell, Point toOffset);


        Boolean RightAngleAxes { get; set; }
        IXLChart SetRightAngleAxes();
        IXLChart SetRightAngleAxes(Boolean rightAngleAxes);

        XLChartType ChartType { get; set; }
        IXLChart SetChartType(XLChartType chartType);

    }
}
