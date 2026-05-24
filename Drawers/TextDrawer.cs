using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TextBoxContent = DocumentFormat.OpenXml.Wordprocessing.TextBoxContent;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using WPS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;


namespace FastReportToWord.Drawers
{
    public static class TextDrawer
    {
        private static uint _currentId = 1;

        public static Run Draw(
            WordprocessingDocument doc,
            FastReport.ReportPage page,
            FastReport.TextObject txt)
        {
            var body = doc.MainDocumentPart.Document.Body;

            long leftEmu =
                (long)(
                    UnitConverter.MmToEmu(page.LeftMargin) +
                    UnitConverter.PointToEmu(txt.AbsLeft));

            long topEmu =
                (long)(
                    UnitConverter.MmToEmu(page.TopMargin) +
                    UnitConverter.PointToEmu(txt.AbsTop));

            long widthEmu =
                (long)UnitConverter.PointToEmu(txt.Width);

            long heightEmu =
                (long)UnitConverter.PointToEmu(txt.Height);

            //long leftEmu = (long)(UnitConverter.PointToEmu(Helpers.ToPoints(page.LeftMargin)) + UnitConverter.PointToEmu(Helpers.ToPoints(txt.AbsLeft)));
            //long topEmu = (long)(UnitConverter.MmToEmu(Helpers.ToPoints(page.TopMargin)) + UnitConverter.PointToEmu(Helpers.ToPoints(txt.AbsTop)));
            //long widthEmu = (long)UnitConverter.PointToEmu(Helpers.ToPoints(txt.Width));
            //long heightEmu = (long)UnitConverter.PointToEmu(Helpers.ToPoints(txt.Height));

            //var paraProps = new ParagraphProperties(
            //    new SpacingBetweenLines()
            //    {
            //        Before = "0",
            //        After = "0",
            //        Line = "0",
            //        LineRule = LineSpacingRuleValues.Exact
            //    }
            //);

            //var paragraph = new Paragraph(
            //    paraProps,
            //    new Run(
            //        CreateDrawing(txt, leftEmu, topEmu, widthEmu, heightEmu)
            //    )
            //);

            //body.Append(paragraph);
            return new Run(CreateDrawing(txt, leftEmu, topEmu, widthEmu, heightEmu));
        }

        private static Drawing CreateDrawing(
            FastReport.TextObject txt,
            long leftEmu,
            long topEmu,
            long widthEmu,
            long heightEmu)
        {
            var shape = CreateShape(txt, widthEmu, heightEmu);

            var graphic = new A.Graphic(
                new A.GraphicData(shape)
                {
                    Uri =
                        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
                });

            var anchor =
                new DW.Anchor(
                    new DW.SimplePosition()
                    {
                        X = 0,
                        Y = 0
                    },

                    new DW.HorizontalPosition(
                        new DW.PositionOffset(leftEmu.ToString()))
                    {
                        RelativeFrom =
                            DW.HorizontalRelativePositionValues.Page
                    },

                    new DW.VerticalPosition(
                        new DW.PositionOffset(topEmu.ToString()))
                    {
                        RelativeFrom =
                            DW.VerticalRelativePositionValues.Page
                    },

                    new DW.Extent()
                    {
                        Cx = widthEmu,
                        Cy = heightEmu
                    },

                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },

                    new DW.WrapNone(),

                    new DW.DocProperties()
                    {
                        Id = GetNextId(),
                        Name = $"TextBox_{txt.Name}"
                    },

                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks()
                        {
                            NoChangeAspect = true
                        }),

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

            return new Drawing(anchor);
        }

        private static WPS.WordprocessingShape CreateShape(
    FastReport.TextObject txt,
    long widthEmu,
    long heightEmu)
        {
            var shape = new WPS.WordprocessingShape();

            shape.Append(
                new WPS.NonVisualDrawingProperties()
                {
                    Id = GetNextId(),
                    Name = $"Shape_{txt.Name}"
                });

            shape.Append(
                new WPS.NonVisualDrawingShapeProperties(
                    new A.ShapeLocks() { NoTextEdit = false }
                ));

            shape.Append(
                CreateShapeProperties(txt, widthEmu, heightEmu));

            shape.Append(new WPS.ShapeStyle(
                new A.LineReference(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }) { Index = 0 },
                new A.FillReference(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }) { Index = 0 },
                new A.EffectReference(new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 }) { Index = 0 },
                new A.FontReference(new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 })
                {
                    Index = A.FontCollectionIndexValues.Minor
                }
            ));

            // Wrap TextBoxContent inside TextBoxInfo2
            shape.Append(
                new WPS.TextBoxInfo2(
                    CreateTextBoxContent(txt)));

            // TextBodyProperties must come AFTER TextBoxInfo2
            shape.Append(
                CreateTextBodyProperties(txt));

            return shape;
        }

        private static WPS.ShapeProperties CreateShapeProperties(
            FastReport.TextObject txt,
            long widthEmu,
            long heightEmu)
        {
            var props = new WPS.ShapeProperties();

            props.Append(
                new A.Transform2D(
                    new A.Offset()
                    {
                        X = 0L,
                        Y = 0L
                    },
                    new A.Extents()
                    {
                        Cx = widthEmu,
                        Cy = heightEmu
                    }));

            props.Append(
                new A.PresetGeometry(
                    new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle
                });

            // fill
            if (txt.FillColor != System.Drawing.Color.Transparent)
            {
                props.Append(
                    new A.SolidFill(
                        new A.RgbColorModelHex()
                        {
                            Val = Helpers.ColorToHex(txt.FillColor)
                        }));
            }
            else
            {
                props.Append(new A.NoFill());
            }

            // border
            if (txt.Border.Lines != FastReport.BorderLines.None)
            {
                props.Append(CreateOutline(txt.Border));
            }
            else
            {
                props.Append(new A.Outline(new A.NoFill()));
            }

            return props;
        }

        private static A.Outline CreateOutline(FastReport.Border border)
        {
            int width =
                (int)UnitConverter.PointToEmu(Helpers.ToPoints(border.Width));

            var outline =
                new A.Outline()
                {
                    Width = width,
                    CapType = A.LineCapValues.Round
                };

            outline.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex()
                    {
                        Val = Helpers.ColorToHex(border.Color)
                    }));

            return outline;
        }

        private static WPS.TextBodyProperties CreateTextBodyProperties(
            FastReport.TextObject txt)
        {
            var bodyProps = new WPS.TextBodyProperties();

            bodyProps.LeftInset =
                (int)Helpers.ToEMU(txt.Padding.Left);

            bodyProps.TopInset =
                (int)Helpers.ToEMU(txt.Padding.Top);

            bodyProps.RightInset =
                (int)Helpers.ToEMU(txt.Padding.Right);

            bodyProps.BottomInset =
                (int)Helpers.ToEMU(txt.Padding.Bottom);

            bodyProps.Wrap =
                txt.WordWrap
                    ? A.TextWrappingValues.Square
                    : A.TextWrappingValues.None;

            bodyProps.Anchor =
                txt.VertAlign switch
                {
                    FastReport.VertAlign.Center =>
                        A.TextAnchoringTypeValues.Center,

                    FastReport.VertAlign.Bottom =>
                        A.TextAnchoringTypeValues.Bottom,

                    _ =>
                        A.TextAnchoringTypeValues.Top
                };

            bodyProps.AnchorCenter = false;

            bodyProps.Append(new A.NoAutoFit());

            return bodyProps;
        }

        private static TextBoxContent CreateTextBoxContent(
            FastReport.TextObject txt)
        {
            var content = new TextBoxContent();

            var paragraph = new Paragraph();

            var paraProps = new ParagraphProperties();

            paraProps.Append(
                new Justification()
                {
                    Val =
                        txt.HorzAlign switch
                        {
                            FastReport.HorzAlign.Center =>
                                JustificationValues.Center,

                            FastReport.HorzAlign.Right =>
                                JustificationValues.Right,

                            FastReport.HorzAlign.Justify =>
                                JustificationValues.Both,

                            _ =>
                                JustificationValues.Left
                        }
                });

            if (!string.IsNullOrEmpty(txt.Text) && Helpers.IsRtlText(txt.Text))
            {
                paraProps.Append(new BiDi());
            }

            paragraph.Append(paraProps);

            var run = new Run();

            run.Append(CreateRunProperties(txt));

            run.Append(
                new Text(txt.Text ?? string.Empty)
                {
                    Space =
                        SpaceProcessingModeValues.Preserve
                });

            paragraph.Append(run);

            content.Append(paragraph);

            return content;
        }

        private static RunProperties CreateRunProperties(
            FastReport.TextObject txt)
        {
            var props = new RunProperties();

            props.Append(
                new RunFonts()
                {
                    Ascii = txt.Font.Name,
                    HighAnsi = txt.Font.Name,
                    ComplexScript = txt.Font.Name,
                    EastAsia = txt.Font.Name,
                });

            string size =
                ((int)(txt.Font.Size * 2)).ToString();

            props.Append(
                new FontSize()
                {
                    Val = size
                });

            props.Append(
                new FontSizeComplexScript()
                {
                    Val = size
                });

            props.Append(
                new Color()
                {
                    Val = Helpers.ColorToHex(txt.TextColor)
                });

            if (txt.Font.Bold)
            {
                props.Append(new Bold());
                props.Append(new BoldComplexScript());
            }

            if (txt.Font.Italic)
            {
                props.Append(new Italic());
                props.Append(new ItalicComplexScript());
            }

            if (txt.Font.Underline)
            {
                props.Append(
                    new Underline()
                    {
                        Val = UnderlineValues.Single
                    });
            }

            if (txt.Font.Strikeout)
            {
                props.Append(new Strike());
            }

            return props;
        }

        private static UInt32Value GetNextId()
        {
            return new UInt32Value(_currentId++);
        }
    }
}
