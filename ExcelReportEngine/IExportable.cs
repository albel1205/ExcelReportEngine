using System.IO;

namespace ExcelReportEngine
{
    public interface IExportable
    {
        MemoryStream GetReportStream();
    }
}
