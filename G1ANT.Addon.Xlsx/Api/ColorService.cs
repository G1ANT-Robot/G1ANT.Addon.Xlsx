using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace G1ANT.Addon.Xlsx.Api
{
    public class ColorService
    {
        private readonly SpreadsheetDocument document;

        private static Dictionary<int, System.Drawing.Color> indexedColors;
        private static System.Drawing.Color[] themeColors;

        public ColorService(SpreadsheetDocument document)
        {
            this.document = document;
        }

        public System.Drawing.Color? GetCellBackgroundColor(Cell cell)
        {
            var cellStyleIndex = GetCellStyleIndex(cell);

            var styles = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
            var cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];
            var fill = (Fill)styles.Stylesheet.Fills.ChildElements[(int)cellFormat.FillId.Value];

            var pf = fill.PatternFill;

            if (pf.PatternType == PatternValues.None)
            {
                return null;
            }

            return GetColor(pf.ForegroundColor);
        }

        public void SetCellBackgroundColor(Cell cell, System.Drawing.Color? rgb)
        {
            var styleIndex = GetCellStyleIndex(cell);
            var styles = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();

            if (rgb == null)
            {
                var newCellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[styleIndex].Clone();
                newCellFormat.FillId = 0;

                var formatIndex = styles.Stylesheet.CellFormats.Count;
                styles.Stylesheet.CellFormats.Append(newCellFormat);
                cell.StyleIndex = formatIndex;
            }
            else
            {
                var targetColor = rgb.Value.A.ToString("X2") + rgb.Value.R.ToString("X2") + rgb.Value.G.ToString("X2") + rgb.Value.B.ToString("X2");

                var newFillId = styles.Stylesheet.Fills.Count;
                var newFormatId = styles.Stylesheet.CellFormats.Count;

                var cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[styleIndex];
                var newCellFormat = (CellFormat)cellFormat.Clone();

                var fill = (Fill)styles.Stylesheet.Fills.ChildElements[(int)newCellFormat.FillId.Value];
                var newFill = (Fill)fill.Clone();

                newCellFormat.FillId = newFillId;

                var newPatternFill = new PatternFill() { PatternType = PatternValues.Solid };
                newPatternFill.ForegroundColor = new ForegroundColor { Rgb = targetColor };
                newPatternFill.BackgroundColor = new BackgroundColor { Indexed = 64U };
                newFill.PatternFill = newPatternFill;

                styles.Stylesheet.Fills.Append(newFill);
                styles.Stylesheet.Fills.Count = new DocumentFormat.OpenXml.UInt32Value((uint)styles.Stylesheet.Fills.ChildElements.Count);

                styles.Stylesheet.CellFormats.Append(newCellFormat);
                styles.Stylesheet.CellFormats.Count = new DocumentFormat.OpenXml.UInt32Value((uint)styles.Stylesheet.CellFormats.ChildElements.Count);

                cell.StyleIndex = newFormatId;
            }
        }

        public System.Drawing.Color? GetCellFontColor(Cell cell)
        {
            var cellStyleIndex = GetCellStyleIndex(cell);

            var styles = document.WorkbookPart.GetPartsOfType<WorkbookStylesPart>().First();
            var cellFormat = (CellFormat)styles.Stylesheet.CellFormats.ChildElements[cellStyleIndex];
            var font = (Font)styles.Stylesheet.Fonts.ChildElements[(int)cellFormat.FontId.Value];

            return GetColor(font.Color);
        }

        public void SetCellFontColor(Cell cell, System.Drawing.Color? rgb)
        {
            var targetColor = rgb.Value.A.ToString("X2") + rgb.Value.R.ToString("X2") + rgb.Value.G.ToString("X2") + rgb.Value.B.ToString("X2");

            var styles = document.WorkbookPart.WorkbookStylesPart;
            var stylesheet = styles.Stylesheet;
            var cellformats = stylesheet.CellFormats;
            var fonts = stylesheet.Fonts;

            var fontIndex = fonts.Count;
            var formatIndex = cellformats.Count;

            var format = (CellFormat)cellformats.ElementAt((int)cell.StyleIndex.Value);

            var font = (Font)fonts.ElementAt((int)format.FontId.Value);
            var newfont = (Font)font.Clone();
            newfont.Color = new Color() { Rgb = targetColor };
            fonts.Append(newfont);
            fonts.Count = new DocumentFormat.OpenXml.UInt32Value((uint)fonts.ChildElements.Count);

            var newformat = (CellFormat)format.Clone();
            newformat.FontId = fontIndex;
            cellformats.Append(newformat);
            cellformats.Count = new DocumentFormat.OpenXml.UInt32Value((uint)cellformats.ChildElements.Count);
            cell.StyleIndex = formatIndex;
        }

        public System.Drawing.Color[] ThemeColors
        {
            get
            {
                if (themeColors == null)
                {
                    LoadTheme(document);
                }

                return themeColors;
            }
        }

        public static Dictionary<int, System.Drawing.Color> IndexedColors
        {
            get
            {
                if (indexedColors == null)
                {
                    var retVal = new Dictionary<int, System.Drawing.Color>()
                    {
                        {0, System.Drawing.ColorTranslator.FromHtml("#FF000000")},
                        {1, System.Drawing.ColorTranslator.FromHtml("#FFFFFFFF")},
                        {2, System.Drawing.ColorTranslator.FromHtml("#FFFF0000")},
                        {3, System.Drawing.ColorTranslator.FromHtml("#FF00FF00")},
                        {4, System.Drawing.ColorTranslator.FromHtml("#FF0000FF")},
                        {5, System.Drawing.ColorTranslator.FromHtml("#FFFFFF00")},
                        {6, System.Drawing.ColorTranslator.FromHtml("#FFFF00FF")},
                        {7, System.Drawing.ColorTranslator.FromHtml("#FF00FFFF")},
                        {8, System.Drawing.ColorTranslator.FromHtml("#FF000000")},
                        {9, System.Drawing.ColorTranslator.FromHtml("#FFFFFFFF")},
                        {10, System.Drawing.ColorTranslator.FromHtml("#FFFF0000")},
                        {11, System.Drawing.ColorTranslator.FromHtml("#FF00FF00")},
                        {12, System.Drawing.ColorTranslator.FromHtml("#FF0000FF")},
                        {13, System.Drawing.ColorTranslator.FromHtml("#FFFFFF00")},
                        {14, System.Drawing.ColorTranslator.FromHtml("#FFFF00FF")},
                        {15, System.Drawing.ColorTranslator.FromHtml("#FF00FFFF")},
                        {16, System.Drawing.ColorTranslator.FromHtml("#FF800000")},
                        {17, System.Drawing.ColorTranslator.FromHtml("#FF008000")},
                        {18, System.Drawing.ColorTranslator.FromHtml("#FF000080")},
                        {19, System.Drawing.ColorTranslator.FromHtml("#FF808000")},
                        {20, System.Drawing.ColorTranslator.FromHtml("#FF800080")},
                        {21, System.Drawing.ColorTranslator.FromHtml("#FF008080")},
                        {22, System.Drawing.ColorTranslator.FromHtml("#FFC0C0C0")},
                        {23, System.Drawing.ColorTranslator.FromHtml("#FF808080")},
                        {24, System.Drawing.ColorTranslator.FromHtml("#FF9999FF")},
                        {25, System.Drawing.ColorTranslator.FromHtml("#FF993366")},
                        {26, System.Drawing.ColorTranslator.FromHtml("#FFFFFFCC")},
                        {27, System.Drawing.ColorTranslator.FromHtml("#FFCCFFFF")},
                        {28, System.Drawing.ColorTranslator.FromHtml("#FF660066")},
                        {29, System.Drawing.ColorTranslator.FromHtml("#FFFF8080")},
                        {30, System.Drawing.ColorTranslator.FromHtml("#FF0066CC")},
                        {31, System.Drawing.ColorTranslator.FromHtml("#FFCCCCFF")},
                        {32, System.Drawing.ColorTranslator.FromHtml("#FF000080")},
                        {33, System.Drawing.ColorTranslator.FromHtml("#FFFF00FF")},
                        {34, System.Drawing.ColorTranslator.FromHtml("#FFFFFF00")},
                        {35, System.Drawing.ColorTranslator.FromHtml("#FF00FFFF")},
                        {36, System.Drawing.ColorTranslator.FromHtml("#FF800080")},
                        {37, System.Drawing.ColorTranslator.FromHtml("#FF800000")},
                        {38, System.Drawing.ColorTranslator.FromHtml("#FF008080")},
                        {39, System.Drawing.ColorTranslator.FromHtml("#FF0000FF")},
                        {40, System.Drawing.ColorTranslator.FromHtml("#FF00CCFF")},
                        {41, System.Drawing.ColorTranslator.FromHtml("#FFCCFFFF")},
                        {42, System.Drawing.ColorTranslator.FromHtml("#FFCCFFCC")},
                        {43, System.Drawing.ColorTranslator.FromHtml("#FFFFFF99")},
                        {44, System.Drawing.ColorTranslator.FromHtml("#FF99CCFF")},
                        {45, System.Drawing.ColorTranslator.FromHtml("#FFFF99CC")},
                        {46, System.Drawing.ColorTranslator.FromHtml("#FFCC99FF")},
                        {47, System.Drawing.ColorTranslator.FromHtml("#FFFFCC99")},
                        {48, System.Drawing.ColorTranslator.FromHtml("#FF3366FF")},
                        {49, System.Drawing.ColorTranslator.FromHtml("#FF33CCCC")},
                        {50, System.Drawing.ColorTranslator.FromHtml("#FF99CC00")},
                        {51, System.Drawing.ColorTranslator.FromHtml("#FFFFCC00")},
                        {52, System.Drawing.ColorTranslator.FromHtml("#FFFF9900")},
                        {53, System.Drawing.ColorTranslator.FromHtml("#FFFF6600")},
                        {54, System.Drawing.ColorTranslator.FromHtml("#FF666699")},
                        {55, System.Drawing.ColorTranslator.FromHtml("#FF969696")},
                        {56, System.Drawing.ColorTranslator.FromHtml("#FF003366")},
                        {57, System.Drawing.ColorTranslator.FromHtml("#FF339966")},
                        {58, System.Drawing.ColorTranslator.FromHtml("#FF003300")},
                        {59, System.Drawing.ColorTranslator.FromHtml("#FF333300")},
                        {60, System.Drawing.ColorTranslator.FromHtml("#FF993300")},
                        {61, System.Drawing.ColorTranslator.FromHtml("#FF993366")},
                        {62, System.Drawing.ColorTranslator.FromHtml("#FF333399")},
                        {63, System.Drawing.ColorTranslator.FromHtml("#FF333333")},
                        {64, System.Drawing.Color.Transparent}
                    };
                    indexedColors = retVal;
                }
                return indexedColors;
            }
        }        

        private void LoadTheme(SpreadsheetDocument document)
        {
            try
            {
                var uri = new Uri(@"/xl/theme/theme1.xml", UriKind.Relative);
                if (document.Package.PartExists(uri))
                {
                    var part = document.Package.GetPart(uri);
                    var partStream = part.GetStream();

                    var xdoc = XDocument.Load(part.GetStream());
                    var ns = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
                    var themeElements = xdoc
                        .Element(ns + "theme")
                        .Element(ns + "themeElements")
                        .Element(ns + "clrScheme")
                        .Elements()
                        .ToArray();

                    themeColors = new System.Drawing.Color[themeElements.Length];

                    for (int i = 0; i < themeElements.Length; i++)
                    {
                        var rgb = themeElements[i]
                            .Element(ns + "srgbClr")?
                            .Attribute("val")?
                            .Value;

                        if (rgb == null)
                        {
                            rgb = themeElements[i]
                            .Element(ns + "sysClr")?
                            .Attribute("lastClr")?
                            .Value;
                        }

                        if (rgb == null)
                        {
                            rgb = "000000";
                        }

                        ThemeColors[i] = System.Drawing.ColorTranslator.FromHtml("#" + rgb);
                    }
                }
            }
            catch
            {
                throw new Exception("Error reading theme from xlsx file");
            }
        }

        private int GetCellStyleIndex(Cell cell)
        {
            return (int)(cell.StyleIndex?.Value ?? 0);
        }

        private System.Drawing.Color? GetColor(ColorType ct)
        {
            if (ct.Rgb != null)
            {
                return System.Drawing.ColorTranslator.FromHtml($"#{ct.Rgb.Value}");
            }

            if (ct.Indexed != null)
            {
                return IndexedColors[(int)ct.Indexed.Value];
            }

            if (ct.Theme != null)
            {
                return ColorHelper.ApplyTintToRgb(ThemeColors[ct.Theme.Value], ct.Tint?.Value ?? 0);
            }

            return null;
        }
    }
}
