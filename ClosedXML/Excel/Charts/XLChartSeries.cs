#nullable disable

namespace ClosedXML.Excel
{
    public class XLChartSeries : IXLChartSeries
    {

        public string Name { get; set; }

        public string XValues { get; set; }

        public string YValues { get; set; }


        public void Dispose()
        {
            // Leave this empty so that Janitor.Fody can do its work
        }
    }
}
