using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FastReport;
using FastReport.Web;
using FastReportToWord.Drawers;
using System.Drawing;
using System.Windows.Forms;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace FastReportToWord
{
    public class WordGenerator : IWordGenerator
    {
        public byte[] Render(WebReport report)
        {
            using (var stream = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    int pagesCount = report.Report.PreparedPages.Count;

                    for (int i = 0; i < pagesCount; i++)
                    {
                        var reportPage = report.Report.PreparedPages.GetPage(i);
                        RenderPage(doc, reportPage);

                        // اگر صفحه آخر نیست، یه Page Break اضافه کن
                        if (i < pagesCount - 1)
                        {
                            mainPart.Document.Body.Append(
                                new W.Paragraph(
                                    new W.Run(
                                        new W.Break() { Type = BreakValues.Page }
                                    )
                                )
                            );
                        }
                    }

                    mainPart.Document.Save();
                } // اینجا Document بسته میشه و محتوا flush میشه

                return stream.ToArray(); // حالا stream کامله
            }
        }

        public static void SetPageSize(
    WordprocessingDocument doc,
    FastReport.ReportPage page)
        {
            var body = doc.MainDocumentPart.Document.Body;

            // Get or create sectPr
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr == null)
            {
                sectPr = new SectionProperties();
                body.Append(sectPr);
            }

            // FastReport page width/height are in mm
            // Word twips: 1 mm = 56.6929 twips
            int widthTwips = (int)(page.PaperWidth * 56.6929);
            int heightTwips = (int)(page.PaperHeight * 56.6929);

            int marginTopTwips = (int)(page.TopMargin * 56.6929);
            int marginBottomTwips = (int)(page.BottomMargin * 56.6929);
            int marginLeftTwips = (int)(page.LeftMargin * 56.6929);
            int marginRightTwips = (int)(page.RightMargin * 56.6929);

            // Remove existing pgSz and pgMar if any
            sectPr.RemoveAllChildren<PageSize>();
            sectPr.RemoveAllChildren<PageMargin>();

            sectPr.Append(new PageSize()
            {
                Width = (UInt32Value)(uint)widthTwips,
                Height = (UInt32Value)(uint)heightTwips,
                Orient = page.Landscape
                    ? PageOrientationValues.Landscape
                    : PageOrientationValues.Portrait
            });

            sectPr.Append(new PageMargin()
            {
                Top = marginTopTwips,
                Bottom = marginBottomTwips,
                Left = (UInt32Value)(uint)marginLeftTwips,
                Right = (UInt32Value)(uint)marginRightTwips
            });
        }

        // ========== رندر کل صفحه ==========
        public void RenderPage(WordprocessingDocument doc, ReportPage page)
        {
            var body = doc.MainDocumentPart!.Document!.Body!;

            // تنظیم Page Size و Margins
            //float leftMarginMM = page.LeftMargin * 10;
            //float topMarginMM = page.TopMargin * 10;
            //float rightMarginMM = page.RightMargin * 10;
            //float bottomMarginMM = page.BottomMargin * 10;
            //float pageWidthMM = page.PaperWidth * 10;
            //float pageHeightMM = page.PaperHeight * 10;

            //SectionProperties sectionProp1 = new SectionProperties(
            //    new SectionType() { Val = SectionMarkValues.NextPage },
            //    new PageSize()
            //    {
            //        Width = UnitConverter.MmToTwips(page.PaperWidth),
            //        Height = UnitConverter.MmToTwips(page.PaperHeight),
            //        Orient = page.Landscape
            //           ? W.PageOrientationValues.Landscape
            //           : W.PageOrientationValues.Portrait
            //    },
            //    new PageMargin()
            //    {
            //        Left = UnitConverter.MmToTwips(page.LeftMargin),
            //        Right = UnitConverter.MmToTwips(page.RightMargin),
            //        Top = (int)UnitConverter.MmToTwips(page.TopMargin),
            //        Bottom = (int)UnitConverter.MmToTwips(page.BottomMargin)
            //    });

            //body.Append(sectionProp1);


            SetPageSize(doc, page);


            var paragraph = new Paragraph(new ParagraphProperties(
                 new SpacingBetweenLines()
                 {
                     Before = "0",
                     After = "0",
                     Line = "240",
                     LineRule = LineSpacingRuleValues.Auto
                 }
             ));

            // رندر Bandها
            foreach (BandBase band in page.Bands)
            {
                // رسم Background
                //if (band.Fill is SolidFill || band.Border.Lines != BorderLines.None)
                //{
                //    DrawBandBackground(doc, band, leftMarginMM, topMarginMM, pageWidthMM - leftMarginMM - rightMarginMM);
                //}

                // رسم Objects
                foreach (ReportComponentBase obj in band.Objects)
                {

                    if (obj is FastReport.TextObject textObj && obj is not FastReport.Table.TableCell && obj is not FastReport.CellularTextObject)
                    {
                        paragraph.Append(TextDrawer.Draw(doc, page, textObj));
                    }
                    else if (obj is PictureObject pictureObj)
                    {
                        paragraph.Append(PictureDrawer.GetRun(doc, page, pictureObj));
                    }
                    else if (obj is LineObject lineObj)
                    {
                        paragraph.Append(LineDrawer.GetRun(doc, page, lineObj));
                    }
                    else if (obj is FastReport.ReportComponentBase component)
                    {
                        var image = Helpers.RenderFastReportObjectToImage(component);
                        MemoryStream ms = new MemoryStream(image);
                        PictureObject pic = new PictureObject()
                        {
                            Width = component.Width,
                            Height = component.Height,
                            Image = Image.FromStream(ms),
                            Top = component.Top,
                            Left = component.Left,
                            Border = component.Border,
                            Name = component.Name,
                            SizeMode = PictureBoxSizeMode.Zoom
                        };

                        paragraph.Append(PictureDrawer.GetRun(doc, page, pic, component.AbsLeft, component.AbsTop));
                    }
                }
            }
            body.Append(paragraph);
            //DocumentFormat.OpenXml.Validation.OpenXmlValidator validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            //int errorCount = 0;
            //foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo error in validator.Validate(doc))
            //{
            //    errorCount++;
            //    Console.WriteLine($"Error {errorCount}: {error.Description}");
            //    Console.WriteLine($"Path: {error.Path.XPath}");
            //}
        }

        // ========== رسم Background Band ==========

    }
}
