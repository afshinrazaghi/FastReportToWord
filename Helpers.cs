using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastReportToWord
{
    public static class Helpers
    {
        public static long ToEMU(float units)
        {
            // FastReport Units → Points → EMU
            // 1 inch = 96 units = 914400 EMU
            return (long)(units * 9525);
        }


        public static float ToPoints(float value)
        {
            return value * 0.749916457811947f;  // = 0.75f
        }

        public static string ColorToHex(System.Drawing.Color color)
        {
            return $"{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        public static bool IsRtlText(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            // چک کردن اولین کاراکتر برای تشخیص RTL
            var firstChar = text.FirstOrDefault(c => !char.IsWhiteSpace(c));
            if (firstChar == default) return false;

            // Persian/Arabic Unicode Range: 0x0600 - 0x06FF
            return firstChar >= 0x0600 && firstChar <= 0x06FF;
        }

        public static byte[] RenderFastReportObjectToImage(FastReport.ReportComponentBase component, float dpi = 300f)
        {
            // محاسبه ابعاد پیکسلی بر اساس DPI
            int widthPx = (int)Math.Ceiling(component.Width / 96f * dpi);
            int heightPx = (int)Math.Ceiling(component.Height / 96f * dpi);

            if (widthPx <= 0) widthPx = 1;
            if (heightPx <= 0) heightPx = 1;

            using (var bmp = new Bitmap(widthPx, heightPx))
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(System.Drawing.Color.Transparent);

                // مقیاس‌دهی Graphics برای تطابق با DPI بالا
                g.ScaleTransform(dpi / 96f, dpi / 96f);

                // 2. بوم نقاشی را به اندازه مختصات مطلق شیء به عقب می‌کشیم
                // این کار باعث می‌شود شیء هر کجا که در گزارش هست، دقیقاً روی نقطه 0,0 در تصویر ما رسم شود
                g.TranslateTransform(-component.AbsLeft, -component.AbsTop);

                // استفاده از موتور رسم خود FastReport
                var e = new FastReport.Utils.FRPaintEventArgs(g, 1.0f, 1.0f, component.Report.GraphicCache);

                component.Draw(e);


                using (var ms = new MemoryStream())
                {
                    bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    return ms.ToArray();
                }
            }
        }
    }
}
