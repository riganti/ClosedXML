#nullable disable

using System;
using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLCharts: IEnumerable<IXLChart>
    {
        IXLChart Add(IXLChart chart);

        IXLChart Chart(Int32 index);
    }
}
