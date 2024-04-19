#nullable disable
using System;
using DocumentFormat.OpenXml;

namespace ClosedXML.Excel;

internal class XLSeries : IXLSeries
{
    public UInt32 Id { get; set; }
    public UInt32 Order { get; set; }

    public string Name { get; set; }

    public IXLRange XVal { get; set; }

    public IXLRange YVal { get; set; }

    public XLColor OutlineColor { get; set; }
}
