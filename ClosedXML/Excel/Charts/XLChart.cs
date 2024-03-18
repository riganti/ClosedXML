#nullable disable

using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel.Drawings;

namespace ClosedXML.Excel
{
    internal enum XLChartTypeCategory { Bar3D }
    internal enum XLBarOrientation { Vertical, Horizontal }
    internal enum XLBarGrouping { Clustered, Percent, Stacked, Standard }
    internal class XLChart: XLDrawing<IXLChart>, IXLChart
    {
        private Int32 _id;
        private Int32 _height;
        private Int32 _width;

        internal String RelId { get; set; }

        public Int32 Id
        {
            get { return _id; }
            internal set
            {
                if ((Worksheet.Charts.FirstOrDefault(c => c.Id.Equals(value)) ?? this) != this)
                    throw new ArgumentException($"The chart ID '{value}' already exists.");

                _id = value;
            }
        }
        public Int32 Width
        {
            get { return _width; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the width, the placement should be FreeFloating");
                _width = value;
            }
        }

        public Int32 Height
        {
            get { return _height; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the height, the placement should be FreeFloating");
                _height = value;
            }
        }

        public Int32 Left
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Offset.X ?? 0; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(value, this.Top));
            }
        }
        public Int32 Top
        {
            get { return Markers[XLMarkerPosition.TopLeft]?.Offset.Y ?? 0; }
            set
            {
                if (this.Placement != XLPicturePlacement.FreeFloating)
                    throw new ArgumentException("To set the top offset, the placement should be FreeFloating");

                Markers[XLMarkerPosition.TopLeft] = new XLMarker(Worksheet.Cell(1, 1), new Point(this.Left, value));
            }
        }

        public IXLCell BottomRightCell
        {
            get
            {
                return Markers[XLMarkerPosition.BottomRight].Cell;
            }

            private set
            {
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A chart and its anchor cells must be on the same worksheet");

                this.Markers[XLMarkerPosition.BottomRight] = new XLMarker(value);
            }
        }

        public IXLCell TopLeftCell
        {
            get
            {
                return Markers[XLMarkerPosition.TopLeft].Cell;
            }

            private set
            {
                if (!value.Worksheet.Equals(this.Worksheet))
                    throw new InvalidOperationException("A chart and its anchor cells must be on the same worksheet");

                this.Markers[XLMarkerPosition.TopLeft] = new XLMarker(value);
            }
        }

        internal IDictionary<XLMarkerPosition, XLMarker> Markers { get; private set; }

        public XLPicturePlacement Placement { get; set; }

        public IXLWorksheet Worksheet { get; }

        public XLChart(XLWorksheet worksheet)
        {
            Container = this;
            this.Worksheet = worksheet;
            Int32 zOrder;
            if (worksheet.Charts.Any())
                zOrder = worksheet.Charts.Max(c => c.ZOrder) + 1;
            else
                zOrder = 1;
            ZOrder = zOrder;
            ShapeId = worksheet.Workbook.ShapeIdManager.GetNext();
            RightAngleAxes = true;

            this.Placement = XLPicturePlacement.MoveAndSize;
            this.Markers = new Dictionary<XLMarkerPosition, XLMarker>()
            {
                [XLMarkerPosition.TopLeft] = null,
                [XLMarkerPosition.BottomRight] = null
            };
        }

        public Point GetOffset(XLMarkerPosition position)
        {
            return Markers[position].Offset;
        }

        public IXLChart MoveTo(Int32 left, Int32 top)
        {
            this.Placement = XLPicturePlacement.FreeFloating;
            this.Left = left;
            this.Top = top;
            return this;
        }

        public IXLChart MoveTo(IXLCell cell)
        {
            return MoveTo(cell, 0, 0);
        }

        public IXLChart MoveTo(IXLCell cell, Int32 xOffset, Int32 yOffset)
        {
            return MoveTo(cell, new Point(xOffset, yOffset));
        }

        public IXLChart MoveTo(IXLCell cell, Point offset)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));
            this.Placement = XLPicturePlacement.Move;
            this.TopLeftCell = cell;
            this.Markers[XLMarkerPosition.TopLeft].Offset = offset;
            return this;
        }

        public IXLChart MoveTo(IXLCell fromCell, IXLCell toCell)
        {
            return MoveTo(fromCell, 0, 0, toCell, 0, 0);
        }

        public IXLChart MoveTo(IXLCell fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXLCell toCell, Int32 toCellXOffset, Int32 toCellYOffset)
        {
            return MoveTo(fromCell, new Point(fromCellXOffset, fromCellYOffset), toCell, new Point(toCellXOffset, toCellYOffset));
        }

        public IXLChart MoveTo(IXLCell fromCell, Point fromOffset, IXLCell toCell, Point toOffset)
        {
            if (fromCell == null) throw new ArgumentNullException(nameof(fromCell));
            if (toCell == null) throw new ArgumentNullException(nameof(toCell));
            this.Placement = XLPicturePlacement.MoveAndSize;

            this.TopLeftCell = fromCell;
            this.Markers[XLMarkerPosition.TopLeft].Offset = fromOffset;

            this.BottomRightCell = toCell;
            this.Markers[XLMarkerPosition.BottomRight].Offset = toOffset;

            return this;
        }

        public Boolean RightAngleAxes { get; set; }
        public IXLChart SetRightAngleAxes()
        {
            RightAngleAxes = true;
            return this;
        }
        public IXLChart SetRightAngleAxes(Boolean rightAngleAxes)
        {
            RightAngleAxes = rightAngleAxes;
            return this;
        }

        public XLChartType ChartType { get; set; }
        public IXLChart SetChartType(XLChartType chartType)
        {
            ChartType = chartType;
            return this;
        }

        public XLChartTypeCategory ChartTypeCategory
        {
            get
            {
                if (Bar3DCharts.Contains(ChartType))
                    return XLChartTypeCategory.Bar3D;
                else
                    throw new NotImplementedException();

            }
        }

        private HashSet<XLChartType> Bar3DCharts = new HashSet<XLChartType> { 
            XLChartType.BarClustered3D, 
            XLChartType.BarStacked100Percent3D, 
            XLChartType.BarStacked3D, 
            XLChartType.Column3D, 
            XLChartType.ColumnClustered3D, 
            XLChartType.ColumnStacked100Percent3D, 
            XLChartType.ColumnStacked3D
        };

        public XLBarOrientation BarOrientation
        {
            get
            {
                if (HorizontalCharts.Contains(ChartType))
                    return XLBarOrientation.Horizontal;
                else
                    return XLBarOrientation.Vertical;
            }
        }

        private HashSet<XLChartType> HorizontalCharts = new HashSet<XLChartType>{
            XLChartType.BarClustered, 
            XLChartType.BarClustered3D, 
            XLChartType.BarStacked, 
            XLChartType.BarStacked100Percent, 
            XLChartType.BarStacked100Percent3D, 
            XLChartType.BarStacked3D, 
            XLChartType.ConeHorizontalClustered, 
            XLChartType.ConeHorizontalStacked, 
            XLChartType.ConeHorizontalStacked100Percent, 
            XLChartType.CylinderHorizontalClustered, 
            XLChartType.CylinderHorizontalStacked, 
            XLChartType.CylinderHorizontalStacked100Percent, 
            XLChartType.PyramidHorizontalClustered, 
            XLChartType.PyramidHorizontalStacked, 
            XLChartType.PyramidHorizontalStacked100Percent
        };

        public XLBarGrouping BarGrouping
        {
            get
            {
                if (ClusteredCharts.Contains(ChartType))
                    return XLBarGrouping.Clustered;
                else if (PercentCharts.Contains(ChartType))
                    return XLBarGrouping.Percent;
                else if (StackedCharts.Contains(ChartType))
                    return XLBarGrouping.Stacked;
                else
                    return XLBarGrouping.Standard;
            }
        }

        public List<IXLSeries> Series { get; set; } = new List<IXLSeries>();


        public HashSet<XLChartType> ClusteredCharts = new HashSet<XLChartType>()
        {
            XLChartType.BarClustered,
            XLChartType.BarClustered3D,
            XLChartType.ColumnClustered,
            XLChartType.ColumnClustered3D,
            XLChartType.ConeClustered,
            XLChartType.ConeHorizontalClustered,
            XLChartType.CylinderClustered,
            XLChartType.CylinderHorizontalClustered,
            XLChartType.PyramidClustered,
            XLChartType.PyramidHorizontalClustered
        };

        public HashSet<XLChartType> PercentCharts = new HashSet<XLChartType>() { 
            XLChartType.AreaStacked100Percent,
            XLChartType.AreaStacked100Percent3D,
            XLChartType.BarStacked100Percent,
            XLChartType.BarStacked100Percent3D,
            XLChartType.ColumnStacked100Percent,
            XLChartType.ColumnStacked100Percent3D,
            XLChartType.ConeHorizontalStacked100Percent,
            XLChartType.ConeStacked100Percent,
            XLChartType.CylinderHorizontalStacked100Percent,
            XLChartType.CylinderStacked100Percent,
            XLChartType.LineStacked100Percent,
            XLChartType.LineWithMarkersStacked100Percent,
            XLChartType.PyramidHorizontalStacked100Percent,
            XLChartType.PyramidStacked100Percent
        };

        public HashSet<XLChartType> StackedCharts = new HashSet<XLChartType>()
        {
            XLChartType.AreaStacked,
            XLChartType.AreaStacked3D,
            XLChartType.BarStacked,
            XLChartType.BarStacked3D,
            XLChartType.ColumnStacked,
            XLChartType.ColumnStacked3D,
            XLChartType.ConeHorizontalStacked,
            XLChartType.ConeStacked,
            XLChartType.CylinderHorizontalStacked,
            XLChartType.CylinderStacked,
            XLChartType.LineStacked,
            XLChartType.LineWithMarkersStacked,
            XLChartType.PyramidHorizontalStacked,
            XLChartType.PyramidStacked
        };
    }
}
