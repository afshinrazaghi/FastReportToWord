using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastReportToWord
{
    public static class UnitConverter
    {
        private const double MM_TO_TWIPS = 56.692913386;
        private const double POINT_TO_INCH = 1.0 / 72.0; // 1 inch = 72 points
        private const double INCH_TO_EMU = 914400.0;
        private const double INCH_TO_TWIPS = 1440.0;
        // تبدیل mm به Twips (برای Page Size و Margins)
        public static uint MmToTwips(double mm)
        {
            return (uint)Math.Round(mm * MM_TO_TWIPS);
        }

        // تبدیل mm به EMU (برای موقعیت Shapes)
        public static long MmToEmu(double mm)
        {
            // 1 inch = 914400 EMU
            // 1 inch = 25.4 mm
            // پس 1 mm = 914400 / 25.4 = 36000 EMU
            // اما این برای Drawing ML هست

            // برای Word Shapes باید از این استفاده کنی:
            return (long)Math.Round(mm * 914400.0 / 25.4);  // ≈ 36000
        }



        // برای موقعیت و اندازه Shapes (EMU)
        //public static long PointToEmu(double points)
        //{
        //    return (long)Math.Round(points * POINT_TO_INCH * INCH_TO_EMU);
        //}

        // برای Page Settings (Twips)
        public static uint PointToTwips(double points)
        {
            return (uint)Math.Round(points * POINT_TO_INCH * INCH_TO_TWIPS);
        }


        public static uint ConvertToDxa(float fastReportValue)
        {
            // 1 inch = 96 pixels (FastReport)
            // 1 inch = 1440 DXA (OpenXml)
            return (uint)((fastReportValue / 96.0) * 1440);
        }


        //private const long EmuPerPoint = 12700;

        private const long EmuPerPoint = 9535;

        // ضریب تبدیل پیکسل به Point (بر اساس DPI استاندارد 96)
        private const float PointsPerPixel = 72.0f / 96.0f;

        /// <summary>
        /// تبدیل Point به EMU برای استفاده در OpenXML
        /// </summary>
        public static uint PointToEmu(double point)
        {
            return (uint)Math.Round(point * EmuPerPoint);
        }
    }
}
