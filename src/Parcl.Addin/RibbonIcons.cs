using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace Parcl.Addin
{
    /// <summary>
    /// Draws custom ribbon icons for the Parcl add-in using GDI+.
    /// All icons are 32x32 for large buttons, 16x16 for normal buttons.
    /// </summary>
    internal static class RibbonIcons
    {
        private static readonly Color AccentBlue = Color.FromArgb(60, 140, 220);
        private static readonly Color AccentGreen = Color.FromArgb(80, 180, 80);
        private static readonly Color AccentOrange = Color.FromArgb(230, 150, 50);
        private static readonly Color AccentRed = Color.FromArgb(200, 70, 70);
        private static readonly Color IconGray = Color.FromArgb(90, 90, 90);
        private static readonly Color IconWhite = Color.FromArgb(250, 250, 250);

        private static readonly Dictionary<string, Func<int, Bitmap>> Generators =
            new Dictionary<string, Func<int, Bitmap>>(StringComparer.OrdinalIgnoreCase)
            {
                { "btnEncrypt", DrawEncrypt },
                { "btnDecrypt", DrawDecrypt },
                { "btnSign", DrawSign },
                { "btnCertExchange", DrawCertExchange },
                { "btnCertSelector", DrawCertSelector },
                { "btnLookup", DrawLookup },
                { "btnDashboard", DrawDashboard },
                { "btnOptions", DrawOptions },
            };

        public static Bitmap GetIcon(string buttonId, int size)
        {
            if (Generators.TryGetValue(buttonId, out var generator))
                return generator(size);

            // Fallback: blue square with "?"
            return DrawFallback(size);
        }

        private static Bitmap DrawEncrypt(int size)
        {
            // Closed padlock
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                using (var pen = new Pen(AccentBlue, 2.5f * s))
                using (var fill = new SolidBrush(AccentBlue))
                {
                    // Shackle (arc)
                    g.DrawArc(pen, 9 * s, 4 * s, 14 * s, 14 * s, 180, 180);
                    g.DrawLine(pen, 9 * s, 11 * s, 9 * s, 14 * s);
                    g.DrawLine(pen, 23 * s, 11 * s, 23 * s, 14 * s);

                    // Lock body
                    var body = new RectangleF(6 * s, 14 * s, 20 * s, 14 * s);
                    FillRoundedRect(g, fill, body, 3 * s);

                    // Keyhole
                    using (var white = new SolidBrush(IconWhite))
                    {
                        g.FillEllipse(white, 13.5f * s, 18 * s, 5 * s, 5 * s);
                        g.FillRectangle(white, 15 * s, 22 * s, 2 * s, 4 * s);
                    }
                }
            }
            return bmp;
        }

        private static Bitmap DrawDecrypt(int size)
        {
            // Open padlock
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                using (var pen = new Pen(AccentGreen, 2.5f * s))
                using (var fill = new SolidBrush(AccentGreen))
                {
                    // Open shackle (arc + one leg raised)
                    g.DrawArc(pen, 9 * s, 1 * s, 14 * s, 14 * s, 180, 180);
                    g.DrawLine(pen, 9 * s, 8 * s, 9 * s, 14 * s);
                    // Right leg is raised (open)

                    // Lock body
                    var body = new RectangleF(6 * s, 14 * s, 20 * s, 14 * s);
                    FillRoundedRect(g, fill, body, 3 * s);

                    // Keyhole
                    using (var white = new SolidBrush(IconWhite))
                    {
                        g.FillEllipse(white, 13.5f * s, 18 * s, 5 * s, 5 * s);
                        g.FillRectangle(white, 15 * s, 22 * s, 2 * s, 4 * s);
                    }
                }
            }
            return bmp;
        }

        private static Bitmap DrawSign(int size)
        {
            // Pen / signature
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                using (var pen = new Pen(AccentOrange, 2.5f * s) { LineJoin = LineJoin.Round })
                using (var fill = new SolidBrush(AccentOrange))
                {
                    // Pen body (diagonal)
                    var penPoints = new PointF[]
                    {
                        new PointF(22 * s, 4 * s),
                        new PointF(26 * s, 8 * s),
                        new PointF(12 * s, 22 * s),
                        new PointF(8 * s, 18 * s),
                    };
                    g.FillPolygon(fill, penPoints);

                    // Pen tip
                    var tipPoints = new PointF[]
                    {
                        new PointF(8 * s, 18 * s),
                        new PointF(12 * s, 22 * s),
                        new PointF(6 * s, 26 * s),
                    };
                    g.FillPolygon(fill, tipPoints);

                    // Signature line
                    using (var linePen = new Pen(IconGray, 1.5f * s))
                    {
                        g.DrawLine(linePen, 4 * s, 28 * s, 28 * s, 28 * s);
                    }
                }
            }
            return bmp;
        }

        private static Bitmap DrawCertExchange(int size)
        {
            // Certificate with arrow (send)
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                DrawCertBase(g, s, AccentBlue);

                // Arrow pointing right
                using (var pen = new Pen(AccentGreen, 2.5f * s) { EndCap = LineCap.ArrowAnchor })
                {
                    g.DrawLine(pen, 18 * s, 22 * s, 28 * s, 22 * s);
                }
            }
            return bmp;
        }

        private static Bitmap DrawCertSelector(int size)
        {
            // Certificate with checkmark
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                DrawCertBase(g, s, AccentBlue);

                // Checkmark
                using (var pen = new Pen(AccentGreen, 2.5f * s) { LineJoin = LineJoin.Round })
                {
                    g.DrawLines(pen, new PointF[]
                    {
                        new PointF(18 * s, 22 * s),
                        new PointF(22 * s, 26 * s),
                        new PointF(29 * s, 18 * s),
                    });
                }
            }
            return bmp;
        }

        private static Bitmap DrawLookup(int size)
        {
            // Magnifying glass
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                using (var pen = new Pen(AccentBlue, 2.5f * s))
                {
                    // Glass circle
                    g.DrawEllipse(pen, 4 * s, 2 * s, 16 * s, 16 * s);
                    // Handle
                    g.DrawLine(pen, 18 * s, 16 * s, 26 * s, 24 * s);
                }
            }
            return bmp;
        }

        private static Bitmap DrawDashboard(int size)
        {
            // Grid / dashboard
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                using (var fill = new SolidBrush(AccentBlue))
                {
                    FillRoundedRect(g, fill, new RectangleF(4 * s, 4 * s, 10 * s, 10 * s), 2 * s);
                    FillRoundedRect(g, fill, new RectangleF(18 * s, 4 * s, 10 * s, 10 * s), 2 * s);
                    FillRoundedRect(g, fill, new RectangleF(4 * s, 18 * s, 10 * s, 10 * s), 2 * s);
                    using (var fill2 = new SolidBrush(AccentOrange))
                        FillRoundedRect(g, fill2, new RectangleF(18 * s, 18 * s, 10 * s, 10 * s), 2 * s);
                }
            }
            return bmp;
        }

        private static Bitmap DrawOptions(int size)
        {
            // Gear
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            {
                float s = size / 32f;
                float cx = 16 * s, cy = 16 * s;
                float outer = 12 * s, inner = 8 * s;
                int teeth = 8;

                using (var fill = new SolidBrush(IconGray))
                {
                    var path = new GraphicsPath();
                    var points = new PointF[teeth * 4];
                    for (int i = 0; i < teeth; i++)
                    {
                        double a0 = (2 * Math.PI * i / teeth) - Math.PI / teeth * 0.4;
                        double a1 = (2 * Math.PI * i / teeth) + Math.PI / teeth * 0.4;
                        double a2 = (2 * Math.PI * (i + 0.5) / teeth) - Math.PI / teeth * 0.4;
                        double a3 = (2 * Math.PI * (i + 0.5) / teeth) + Math.PI / teeth * 0.4;

                        points[i * 4] = new PointF(cx + outer * (float)Math.Cos(a0), cy + outer * (float)Math.Sin(a0));
                        points[i * 4 + 1] = new PointF(cx + outer * (float)Math.Cos(a1), cy + outer * (float)Math.Sin(a1));
                        points[i * 4 + 2] = new PointF(cx + inner * (float)Math.Cos(a2), cy + inner * (float)Math.Sin(a2));
                        points[i * 4 + 3] = new PointF(cx + inner * (float)Math.Cos(a3), cy + inner * (float)Math.Sin(a3));
                    }
                    path.AddPolygon(points);
                    g.FillPath(fill, path);

                    // Center hole
                    using (var hole = new SolidBrush(Color.Transparent))
                    {
                        g.CompositingMode = CompositingMode.SourceCopy;
                        g.FillEllipse(new SolidBrush(Color.Transparent),
                            cx - 4 * s, cy - 4 * s, 8 * s, 8 * s);
                        g.CompositingMode = CompositingMode.SourceOver;
                    }
                }
            }
            return bmp;
        }

        private static Bitmap DrawFallback(int size)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Smooth(bmp))
            using (var fill = new SolidBrush(AccentBlue))
            {
                g.FillEllipse(fill, 2, 2, size - 4, size - 4);
            }
            return bmp;
        }

        // Shared helper: draw a certificate document outline
        private static void DrawCertBase(Graphics g, float s, Color color)
        {
            using (var pen = new Pen(color, 2 * s))
            using (var fill = new SolidBrush(Color.FromArgb(40, color)))
            {
                // Document outline
                var doc = new RectangleF(4 * s, 2 * s, 18 * s, 24 * s);
                g.FillRectangle(fill, doc);
                g.DrawRectangle(pen, doc.X, doc.Y, doc.Width, doc.Height);

                // Ribbon/seal circle
                using (var seal = new SolidBrush(color))
                    g.FillEllipse(seal, 9 * s, 6 * s, 8 * s, 8 * s);

                // Text lines
                using (var line = new Pen(color, 1.5f * s))
                {
                    g.DrawLine(line, 8 * s, 18 * s, 18 * s, 18 * s);
                    g.DrawLine(line, 10 * s, 21 * s, 16 * s, 21 * s);
                }
            }
        }

        private static Graphics Smooth(Bitmap bmp)
        {
            var g = Graphics.FromImage(bmp);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);
            return g;
        }

        private static void FillRoundedRect(Graphics g, Brush brush, RectangleF rect, float radius)
        {
            using (var path = new GraphicsPath())
            {
                float d = radius * 2;
                path.AddArc(rect.X, rect.Y, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                g.FillPath(brush, path);
            }
        }
    }
}
