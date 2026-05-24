using FastReport.Web;

namespace FastReportToWord
{
    public interface IWordGenerator
    {
        byte[] Render(WebReport report);
    }
}
