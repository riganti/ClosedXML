#nullable disable
using System;

namespace ClosedXML.Excel;

public interface IXLSeries
{
    UInt32 Id { get; set; }
    UInt32 Order { get; set; }

    string Name { get; set; }

    IXLRange XVal { get; set; }

    IXLRange YVal { get; set; }
}
