#nullable disable

using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public interface IXLCharts: IEnumerable<IXLChart>
    {
        IXLChart Add(IXLChart chart);
    }
}
