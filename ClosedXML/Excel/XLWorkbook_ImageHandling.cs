#nullable disable

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Linq;

using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Cdr = DocumentFormat.OpenXml.Drawing.Charts;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        public static OpenXmlElement GetAnchorFromImageId(DrawingsPart drawingsPart, string relId)
        {
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<Xdr.BlipFill>()
                    .Any(x => x?.Blip?.Embed?.Value.Equals(relId) ?? false)
                );
            return matchingAnchor.FirstOrDefault();
        }

        public static OpenXmlElement GetAnchorFromImageIndex(WorksheetPart worksheetPart, Int32 index)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            var matchingAnchor = drawingsPart.WorksheetDrawing
                .Where(wsdr => wsdr.Descendants<Xdr.NonVisualDrawingProperties>()
                    .Any(x => x.Id.Value.Equals(Convert.ToUInt32(index + 1)))
                );

            return matchingAnchor.FirstOrDefault();
        }

        public static NonVisualDrawingProperties GetPropertiesFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
                return null;

            return anchor
                .Descendants<Xdr.NonVisualDrawingProperties>()
                .FirstOrDefault();
        }

        public static String GetImageRelIdFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
                return null;

            var blipFill = anchor.Descendants<Xdr.BlipFill>().FirstOrDefault();
            return blipFill?.Blip?.Embed?.Value;
        }

        public static String GetChartRelIdFromAnchor(OpenXmlElement anchor)
        {
            if (!IsAllowedAnchor(anchor))
                return null;

            var chart = anchor.Descendants<Cdr.ChartReference>().FirstOrDefault();
            return chart?.Id?.Value;
        }

        private static bool IsAllowedAnchor(OpenXmlElement anchor)
        {
            var allowedAnchorTypes = new Type[] { typeof(AbsoluteAnchor), typeof(OneCellAnchor), typeof(TwoCellAnchor) };
            return (allowedAnchorTypes.Any(t => t == anchor.GetType()));
        }
    }
}
