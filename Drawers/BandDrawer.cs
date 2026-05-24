using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FastReport;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace FastReportToWord.Drawers
{
    public static class BandDrawer
    {
        public static void Draw(WordprocessingDocument doc, ReportPage page, FastReport.BandBase band)
        {
            var body = doc.MainDocumentPart.Document.Body;

            float widthEmu = Helpers.ToEMU(band.Width);
            float heightEmu = Helpers.ToEMU(band.Height);

            // ایجاد Paragraph container
            var paragraph = new DocumentFormat.OpenXml.Drawing.Paragraph();
            var run = new DocumentFormat.OpenXml.Drawing.Run();

            // Layer 1: Border (با استفاده از Shape)
            var borderShape = CreateBorderShape(band.Border, widthEmu, heightEmu);
            run.Append(borderShape);

            // Layer 2: Content (با استفاده از TextBox با Absolute Positioning)
            foreach (FastReport.ReportComponentBase obj in band.Objects)
            {
                if (obj is FastReport.TextObject textObj)
                {
                    var textShape = CreateTextShape(textObj);
                    run.Append(textShape);
                }
                else if (obj is FastReport.PictureObject pictureObj)
                {
                    var pictureShape = CreatePictureShape(doc, pictureObj);
                    run.Append(pictureShape);
                }
                else if (obj is FastReport.ShapeObject shapeObj)
                {
                    var shape = CreateShape(shapeObj);
                    run.Append(shape);
                }
                else if (obj is FastReport.LineObject lineObj)
                {
                    var line = CreateLine(lineObj);
                    run.Append(line);
                }
            }

            paragraph.Append(run);
            body.Append(paragraph);
        }

        // Layer 1: Border Shape
        private static Drawing CreateBorderShape(FastReport.Border border, float widthEmu, float heightEmu)
        {
            var drawing = new Drawing();

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset("0"))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset("0"))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent() { Cx = (long)widthEmu, Cy = (long)heightEmu },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Border" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        CreateBorderShapeProperties(border, widthEmu, heightEmu)
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 251658240U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            return drawing;
        }

        private static PIC.Picture CreateBorderShapeProperties(FastReport.Border border,
            float widthEmu, float heightEmu)
        {
            var picture = new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Rectangle" },
                    new PIC.NonVisualPictureDrawingProperties()
                ),
                new PIC.BlipFill(
                    new A.Blip(),
                    new A.Stretch(new A.FillRectangle())
                ),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 0L, Y = 0L },
                        new A.Extents() { Cx = (long)widthEmu, Cy = (long)heightEmu }
                    ),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            // اعمال Border
            if (border != null)
            {
                var shapeProperties = picture.ShapeProperties;

                // Fill Color
                if (border.Color != System.Drawing.Color.Transparent)
                {
                    shapeProperties.Append(new A.SolidFill(
                        new A.RgbColorModelHex() { Val = ColorToHex(border.Color) }
                    ));
                }
                else
                {
                    shapeProperties.Append(new A.NoFill());
                }

                // Border Lines
                var outline = new A.Outline() { Width = (int)(border.Width * 12700) }; // EMU

                if (border.Color != System.Drawing.Color.Transparent)
                {
                    outline.Append(new A.SolidFill(
                        new A.RgbColorModelHex() { Val = ColorToHex(border.Color) }
                    ));
                }

                // Border Style
                if (border.Style == LineStyle.Dash)
                {
                    outline.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Dash });
                }
                else if (border.Style == LineStyle.Dot)
                {
                    outline.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Dot });
                }

                shapeProperties.Append(outline);
            }

            return picture;
        }

        // Layer 2: Text Shape با Absolute Positioning
        private static Drawing CreateTextShape(FastReport.TextObject textObj)
        {
            long leftEmu = Helpers.ToEMU(textObj.Left);
            long topEmu = Helpers.ToEMU(textObj.Top);
            long widthEmu = Helpers.ToEMU(textObj.Width);
            long heightEmu = Helpers.ToEMU(textObj.Height);

            var drawing = new Drawing();

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset(leftEmu.ToString()))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset(topEmu.ToString()))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)2U, Name = "TextBox" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        CreateTextBoxShape(textObj, widthEmu, heightEmu)
                    )
                    { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            return drawing;
        }

        private static OpenXmlElement CreateTextBoxShape(FastReport.TextObject textObj,
            long widthEmu, long heightEmu)
        {
            // ایجاد TextBox با محتوای متنی
            var textBoxInfo = new TextBoxInfo2(
                new TextBoxContent(
                    new DocumentFormat.OpenXml.Drawing.Paragraph(
                        new DocumentFormat.OpenXml.Drawing.ParagraphProperties(
                            new Justification()
                            {
                                Val = ConvertAlignment(textObj.HorzAlign)
                            }
                        ),
                        new DocumentFormat.OpenXml.Drawing.Run(
                            new DocumentFormat.OpenXml.Drawing.RunProperties(
                                new FontSize() { Val = ((int)(textObj.Font.Size * 2)).ToString() },
                                new Bold() { Val = textObj.Font.Bold },
                                new Italic() { Val = textObj.Font.Italic },
                                new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = ColorToHex(textObj.TextColor) },
                                new RunFonts() { Ascii = textObj.Font.Name }
                            ),
                            new DocumentFormat.OpenXml.Drawing.Text(textObj.Text)
                        )
                    )
                )
            );

            return textBoxInfo;
        }

        // Picture Shape
        private static Drawing CreatePictureShape(WordprocessingDocument doc,
            FastReport.PictureObject pictureObj)
        {
            if (pictureObj.Image == null) return null;

            long leftEmu = Helpers.ToEMU(pictureObj.Left);
            long topEmu = Helpers.ToEMU(pictureObj.Top);
            long widthEmu = Helpers.ToEMU(pictureObj.Width);
            long heightEmu = Helpers.ToEMU(pictureObj.Height);

            // افزودن تصویر به document
            var imagePart = doc.MainDocumentPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream())
            {
                pictureObj.Image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                stream.Position = 0;
                imagePart.FeedData(stream);
            }

            var drawing = new Drawing();

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset(leftEmu.ToString()))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset(topEmu.ToString()))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)3U, Name = "Picture" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Image" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip()
                                {
                                    Embed = doc.MainDocumentPart.GetIdOfPart(imagePart)
                                },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            return drawing;
        }

        // Shape Object
        private static Drawing CreateShape(FastReport.ShapeObject shapeObj)
        {
            long leftEmu = Helpers.ToEMU(shapeObj.Left);
            long topEmu = Helpers.ToEMU(shapeObj.Top);
            long widthEmu = Helpers.ToEMU(shapeObj.Width);
            long heightEmu = Helpers.ToEMU(shapeObj.Height);

            var shapeType = ConvertShapeType(shapeObj.Shape);

            var drawing = new Drawing();

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset(leftEmu.ToString()))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset(topEmu.ToString()))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)4U, Name = "Shape" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Shape" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip(),
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = shapeType },
                                new A.SolidFill(
                                    new A.RgbColorModelHex() { Val = ColorToHex(shapeObj.FillColor) }
                                ),
                                new A.Outline() { Width = (int)(shapeObj.Border.Width * 12700) }
                                    .AppendChild(new A.SolidFill(
                                        new A.RgbColorModelHex() { Val = ColorToHex(shapeObj.Border.Color) }
                                    ))
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            return drawing;
        }

        // Line Object
        private static Drawing CreateLine(FastReport.LineObject lineObj)
        {
            long leftEmu = Helpers.ToEMU(lineObj.Left);
            long topEmu = Helpers.ToEMU(lineObj.Top);
            long widthEmu = Helpers.ToEMU(lineObj.Width);
            long heightEmu = Helpers.ToEMU(lineObj.Height);

            var drawing = new Drawing();

            var anchor = new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset(leftEmu.ToString()))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset(topEmu.ToString()))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent() { Cx = widthEmu, Cy = heightEmu },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)5U, Name = "Line" },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties() { Id = 0U, Name = "Line" },
                                new PIC.NonVisualPictureDrawingProperties()
                            ),
                            new PIC.BlipFill(
                                new A.Blip(),
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = lineObj.Diagonal ? A.ShapeTypeValues.Line : A.ShapeTypeValues.StraightConnector1
                                },
                                new A.Outline() { Width = (int)(lineObj.Border.Width * 12700) }
                                    .AppendChild(new A.SolidFill(
                                        new A.RgbColorModelHex() { Val = ColorToHex(lineObj.Border.Color) }
                                    ))
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            drawing.Append(anchor);
            return drawing;
        }

        // Helper Methods
        private static string ColorToHex(System.Drawing.Color color)
        {
            return $"{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private static JustificationValues ConvertAlignment(FastReport.HorzAlign align)
        {
            return align switch
            {
                FastReport.HorzAlign.Left => JustificationValues.Left,
                FastReport.HorzAlign.Center => JustificationValues.Center,
                FastReport.HorzAlign.Right => JustificationValues.Right,
                FastReport.HorzAlign.Justify => JustificationValues.Both,
                _ => JustificationValues.Left
            };
        }

        private static A.ShapeTypeValues ConvertShapeType(FastReport.ShapeKind shape)
        {
            return shape switch
            {
                FastReport.ShapeKind.Rectangle => A.ShapeTypeValues.Rectangle,
                FastReport.ShapeKind.RoundRectangle => A.ShapeTypeValues.RoundRectangle,
                FastReport.ShapeKind.Ellipse => A.ShapeTypeValues.Ellipse,
                FastReport.ShapeKind.Diamond => A.ShapeTypeValues.Diamond,
                FastReport.ShapeKind.Triangle => A.ShapeTypeValues.Triangle,
                _ => A.ShapeTypeValues.Rectangle
            };
        }
    }
}
