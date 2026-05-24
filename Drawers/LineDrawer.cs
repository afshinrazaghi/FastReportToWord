using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FastReport;
using System;
using System.Globalization;
using DW = DocumentFormat.OpenXml.Drawing;
using DWP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace FastReportToWord.Drawers
{
    public class LineDrawer
    {
        public static DocumentFormat.OpenXml.Wordprocessing.Run GetRun(WordprocessingDocument document, ReportPage page, LineObject lineObject)
        {

            // Calculate absolute coordinates in EMUs
            long widthEmu = UnitConverter.PointToEmu(lineObject.Width);
            long heightEmu = UnitConverter.PointToEmu(lineObject.Height);

            long startX = (long)(
                         UnitConverter.MmToEmu(page.LeftMargin) +
                        UnitConverter.PointToEmu(lineObject.AbsLeft));

            long startY = (long)(
                        UnitConverter.MmToEmu(page.TopMargin) +
                        UnitConverter.PointToEmu(lineObject.AbsTop));

            // Calculate end coordinates
            long x1 = startX;
            long y1 = startY;
            long x2 = startX + widthEmu;
            long y2 = startY + heightEmu;

            // FastReport allows negative width/height for diagonal lines, handle coordinates accordingly
            //if (lineObject.Diagonal)
            //{
            //    // If width or height is negative, the End coordinate is actually less than the Start coordinate
            //    if (lineObject.Width < 0) x2 = x1 + (long)(Helpers.ToPoints(lineObject.Width) * ptToEmu);
            //    if (lineObject.Height < 0) y2 = y1 + (long)(Helpers.ToPoints(lineObject.Height) * ptToEmu);
            //}
            //else
            //{
            // Straight line: Force it to be perfectly horizontal or vertical
            //if (Math.Abs(lineObject.Width) > Math.Abs(lineObject.Height))
            //    y2 = y1;
            //else
            //    x2 = x1;
            //}

            // Format coordinates for VML using "emu" suffix
            string fromPoint = $"{x1}emu,{y1}emu";
            string toPoint = $"{x2}emu,{y2}emu";

            // Extract border properties
            string hexColor = $"#{lineObject.Border.Color.R:X2}{lineObject.Border.Color.G:X2}{lineObject.Border.Color.B:X2}";

            // Stroke weight can also be in EMU, or you can leave it in points depending on your preference. 
            // VML handles both, but here it is converted to EMU as requested.
            long weightEmu = (long)UnitConverter.PointToEmu(lineObject.Border.Width);
            string weight = $"{weightEmu}emu";

            // Create the VML Line
            var vmlLine = new DocumentFormat.OpenXml.Vml.Line()
            {
                From = fromPoint,
                To = toPoint,
                StrokeWeight = weight,
                StrokeColor = hexColor,
                // position:absolute is required to float the line exactly where it was in FastReport
                Style = "position:absolute; mso-position-horizontal-relative:page; mso-position-vertical-relative:page; z-index:1;"
            };

            // Handle Line Style (Dashed, Dotted, etc.)
            if (lineObject.Border.Style != FastReport.LineStyle.Solid)
            {
                var stroke = new DocumentFormat.OpenXml.Vml.Stroke();
                switch (lineObject.Border.Style)
                {
                    case FastReport.LineStyle.Dash: stroke.DashStyle = "dash"; break;
                    case FastReport.LineStyle.Dot: stroke.DashStyle = "dot"; break;
                    case FastReport.LineStyle.DashDot: stroke.DashStyle = "dashdot"; break;
                    case FastReport.LineStyle.DashDotDot: stroke.DashStyle = "longdashdotdot"; break;
                }
                vmlLine.Append(stroke);
            }

            // Wrap in Picture and Run
            var picture = new DocumentFormat.OpenXml.Wordprocessing.Picture(vmlLine);
            var run = new DocumentFormat.OpenXml.Wordprocessing.Run(picture);
            return run;
            
            //var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

            // Safe insertion (checking for sectPr to avoid schema errors as discussed previously)
            //var body = document.MainDocumentPart.Document.Body;
            //var sectPr = body.Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>().LastOrDefault();
            //if (sectPr != null)
            //{
            //    body.InsertBefore(paragraph, sectPr);
            //}
            //else
            //{
            //    body.AppendChild(paragraph);
            //}

        }


    }

}
