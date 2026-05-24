using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

using System;
using System.Drawing;
using System.IO;

namespace FastReportToWord.Drawers
{
    public static class PictureDrawer
    {
        public static Run GetRun(
            WordprocessingDocument doc,
            FastReport.ReportPage page,
            FastReport.PictureObject pic,
            float absLeft = 0,
            float absTop = 0)
        {
            if (pic.Image == null)
                pic.ForceLoadImage();

            if (pic.Image == null)
                return new Run();

            absLeft = absLeft > 0 ? absLeft : pic.AbsLeft;
            absTop = absTop > 0 ? absTop : pic.AbsTop;

            var body = doc.MainDocumentPart.Document.Body;

            long leftEmu =
                (long)(
                    UnitConverter.MmToEmu(page.LeftMargin) +
                    UnitConverter.PointToEmu(absLeft));

            long topEmu =
                (long)(
                    UnitConverter.MmToEmu(page.TopMargin) +
                    UnitConverter.PointToEmu(absTop));

            long widthEmu = (long)UnitConverter.PointToEmu(pic.Width);
            long heightEmu = (long)UnitConverter.PointToEmu(pic.Height);

            long borderEmu =
                (pic.Border != null &&
                 pic.Border.Lines != FastReport.BorderLines.None)
                    ? (long)UnitConverter.PointToEmu(pic.Border.Width)
                    : 0;

            double imgW = pic.Image.Width;
            double imgH = pic.Image.Height;
            double boxRatio = (double)widthEmu / heightEmu;
            double imgRatio = imgW / imgH;

            long finalWidth = widthEmu;
            long finalHeight = heightEmu;

            if (pic.SizeMode == System.Windows.Forms.PictureBoxSizeMode.Zoom)
            {
                if (imgRatio > boxRatio)
                    finalHeight = (long)(widthEmu / imgRatio);
                else
                    finalWidth = (long)(heightEmu * imgRatio);
            }
            else if (pic.SizeMode == System.Windows.Forms.PictureBoxSizeMode.CenterImage)
            {
                long imgEmuW = (long)UnitConverter.PointToEmu((float)imgW);
                long imgEmuH = (long)UnitConverter.PointToEmu((float)imgH);

                finalWidth = imgEmuW;
                finalHeight = imgEmuH;

                leftEmu += (widthEmu - finalWidth) / 2;
                topEmu += (heightEmu - finalHeight) / 2;
            }

            leftEmu += borderEmu / 2;
            topEmu += borderEmu / 2;

            string relationshipId = AddImagePart(doc, pic.Image);

            // Use shared ID counter to avoid conflicts with TextDrawer
            uint docPropsId = DrawingIdCounter.Next();
            uint picId = DrawingIdCounter.Next();

            var shapeProps = new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = 0, Y = 0 },
                    new A.Extents() { Cx = finalWidth, Cy = finalHeight }
                ),
                new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle
                }
            );

            if (pic.Border != null &&
                pic.Border.Lines != FastReport.BorderLines.None)
            {
                shapeProps.Append(
                    new A.Outline(
                        new A.SolidFill(
                            new A.RgbColorModelHex()
                            {
                                Val = Helpers.ColorToHex(pic.Border.Color)
                            }))
                    {
                        Width = (int)UnitConverter.PointToEmu(pic.Border.Width),
                        CapType = A.LineCapValues.Flat,
                        CompoundLineType = A.CompoundLineValues.Single
                    });
            }
            else
            {
                shapeProps.Append(new A.Outline(new A.NoFill()));
            }

            var picture = new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties()
                    {
                        Id = new UInt32Value(picId),
                        Name = $"Picture_{pic.Name}"
                    },
                    new PIC.NonVisualPictureDrawingProperties()
                ),
                new PIC.BlipFill(
                    new A.Blip()
                    {
                        Embed = relationshipId,
                        CompressionState = A.BlipCompressionValues.Print
                    },
                    new A.Stretch(new A.FillRectangle())
                ),
                shapeProps
            );

            var graphic = new A.Graphic(
                new A.GraphicData(picture)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                });

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0, Y = 0 },

                new DW.HorizontalPosition(
                    new DW.PositionOffset(leftEmu.ToString()))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },

                new DW.VerticalPosition(
                    new DW.PositionOffset(topEmu.ToString()))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },

                new DW.Extent() { Cx = finalWidth, Cy = finalHeight },

                new DW.EffectExtent()
                {
                    LeftEdge = 0,
                    TopEdge = 0,
                    RightEdge = 0,
                    BottomEdge = 0
                },

                new DW.WrapNone(),

                new DW.DocProperties()
                {
                    Id = new UInt32Value(docPropsId),
                    Name = $"Picture_{pic.Name}"
                },

                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),

                graphic
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                RelativeHeight = 251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true,
                SimplePos = false
            };

            return new Run(new Drawing(anchor));

            //body.Append(new Paragraph(new Run(new Drawing(anchor))));
        }

        private static string AddImagePart(
            WordprocessingDocument doc,
            Image image)
        {
            var imagePart =
                doc.MainDocumentPart.AddImagePart(ImagePartType.Png);

            using (var ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                imagePart.FeedData(ms);
            }

            return doc.MainDocumentPart.GetIdOfPart(imagePart);
        }
    }

    public static class DrawingIdCounter
    {
        private static uint _current = 1;

        public static uint Next()
        {
            return _current++;
        }

        public static void Reset()
        {
            _current = 1;
        }
    }
}
