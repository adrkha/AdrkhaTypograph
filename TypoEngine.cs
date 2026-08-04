using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Windows.Media.Imaging;
using HarfBuzzSharp;
using SkiaSharp;

namespace AdrkhaTypograph
{
    public enum TypoAlignment
    {
        Right,
        Center,
        Left
    }

    public class TypoEngine
    {
        private struct GlyphInfo
        {
            public ushort Codepoint;
            public float X;
            public float Y;
            public SKPath Path;
            public SKRect TightBounds;
        }

        private struct GlyphInfoWrapper
        {
            public ushort GlyphId;
            public int Cluster;
        }

        public static Dictionary<ushort, int> BuildGlyphToUnicodeMap(SKFont font)
        {
            var map = new Dictionary<ushort, int>();
            int[][] ranges = new int[][]
            {
                new int[] { 0x0020, 0x007F }, // Basic Latin
                new int[] { 0x0600, 0x06FF }, // Arabic
                new int[] { 0x0750, 0x077F }, // Arabic Supplement
                new int[] { 0x0870, 0x08FF }, // Arabic Extended-A & B
                new int[] { 0xE000, 0xF8FF }, // Private Use Area (PUA)
                new int[] { 0xFB50, 0xFDFF }, // Arabic Presentation Forms-A
                new int[] { 0xFE70, 0xFEFF }  // Arabic Presentation Forms-B
            };

            foreach (var range in ranges)
            {
                int start = range[0];
                int end = range[1];
                for (int cp = start; cp <= end; cp++)
                {
                    try
                    {
                        string s = char.ConvertFromUtf32(cp);
                        ushort[] g = font.GetGlyphs(s);
                        if (g != null && g.Length > 0 && g[0] != 0)
                        {
                            if (!map.ContainsKey(g[0]))
                                map[g[0]] = cp;
                        }
                    }
                    catch {}
                }
            }
            return map;
        }

        public static string GetShapedTextString(string text, string fontPath, List<string> activeFeatures)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
                return text;

            try
            {
                using (var fs1 = File.OpenRead(fontPath))
                using (var fs2 = File.OpenRead(fontPath))
                using (var typeface = SKTypeface.FromStream(fs1))
                using (var font = new SKFont(typeface, 24))
                using (var blob = Blob.FromStream(fs2))
                using (var face = new Face(blob, 0))
                using (var hbFont = new HarfBuzzSharp.Font(face))
                {
                    hbFont.SetFunctionsOpenType();
                    hbFont.SetScale(face.UnitsPerEm, face.UnitsPerEm);

                    var glyphToUnicode = BuildGlyphToUnicodeMap(font);

                    string[] textLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    var shapedLines = new List<string>();

                    for (int j = 0; j < textLines.Length; j++)
                    {
                        string lineText = textLines[j];
                        if (string.IsNullOrEmpty(lineText))
                        {
                            shapedLines.Add(string.Empty);
                            continue;
                        }

                        using (var buffer = new HarfBuzzSharp.Buffer())
                        {
                            buffer.Direction = Direction.RightToLeft;
                            buffer.Script = Script.Arabic;
                            buffer.Language = new Language("ar");
                            buffer.AddUtf8(lineText);
                            buffer.GuessSegmentProperties();

                            var features = new List<Feature>();
                            if (activeFeatures != null)
                            {
                                foreach (var f in activeFeatures)
                                    features.Add(new Feature(Tag.Parse(f), 1, 0, uint.MaxValue));
                            }

                            hbFont.Shape(buffer, features.ToArray());

                            var glyphInfos = buffer.GlyphInfos;
                            var sb = new StringBuilder();

                            var sortedGlyphs = new List<GlyphInfoWrapper>();
                            for (int i = 0; i < glyphInfos.Length; i++)
                            {
                                sortedGlyphs.Add(new GlyphInfoWrapper
                                {
                                    GlyphId = (ushort)glyphInfos[i].Codepoint,
                                    Cluster = (int)glyphInfos[i].Cluster
                                });
                            }

                            sortedGlyphs.Sort((g1, g2) => g1.Cluster.CompareTo(g2.Cluster));

                            foreach (var g in sortedGlyphs)
                            {
                                if (glyphToUnicode.TryGetValue(g.GlyphId, out int unicodeChar))
                                {
                                    sb.Append(char.ConvertFromUtf32(unicodeChar));
                                }
                                else
                                {
                                    if (g.Cluster >= 0 && g.Cluster < lineText.Length)
                                    {
                                        if (char.IsHighSurrogate(lineText[g.Cluster]) && g.Cluster + 1 < lineText.Length)
                                        {
                                            sb.Append(lineText.Substring(g.Cluster, 2));
                                        }
                                        else
                                        {
                                            sb.Append(lineText[g.Cluster]);
                                        }
                                    }
                                }
                            }

                            shapedLines.Add(sb.ToString());
                        }
                    }

                    return string.Join(Environment.NewLine, shapedLines);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetShapedTextString error: " + ex.Message);
                return text;
            }
        }

        private class LineInfo
        {
            public List<GlyphInfo> Glyphs = new List<GlyphInfo>();
            public float MinX = float.MaxValue;
            public float MaxX = float.MinValue;
            public float MinY = float.MaxValue;
            public float MaxY = float.MinValue;
            public float Width => MaxX > MinX ? MaxX - MinX : 0f;
        }

        // Create SVG from glyph outlines (vector paths)
        public static string CreateVectorText(string text, string fontPath, List<string> activeFeatures, TypoAlignment alignment)
        {
            string svgPath = Path.Combine(Path.GetTempPath(), "ArabicVector.svg");

            try
            {
                if (string.IsNullOrEmpty(text))
                {
                    File.WriteAllText(svgPath, "<svg xmlns='http://www.w3.org/2000/svg'></svg>");
                    return svgPath;
                }

                // فتح الملف كجداول بيانات ثنائية (Stream) لتفادي مشاكل الحروف العربية (Unicode) في مسارات الملفات بنظام ويندوز
                using (var fs1 = File.OpenRead(fontPath))
                using (var fs2 = File.OpenRead(fontPath))
                using (var typeface = SKTypeface.FromStream(fs1))
                using (var blob = Blob.FromStream(fs2))
                using (var face = new Face(blob, 0))
                using (var font = new HarfBuzzSharp.Font(face))
                {
                    // تفعيل دوال محرك OpenType وتحديد مقاييس الخط ليقوم HarfBuzz بتشكيل الحروف بشكل صحيح
                    font.SetFunctionsOpenType();
                    font.SetScale(face.UnitsPerEm, face.UnitsPerEm);

                    float fontSize = 150f;
                    var skFont = new SKFont(typeface, fontSize);
                    float scale = fontSize / face.UnitsPerEm;
                    var metrics = skFont.Metrics;

                    // تباعد الأسطر الرأسي (1.3x من ارتفاع الخط)
                    float lineHeight = (metrics.Descent - metrics.Ascent) * 1.3f;

                    string[] textLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    var processedLines = new List<LineInfo>();

                    for (int j = 0; j < textLines.Length; j++)
                    {
                        string lineText = textLines[j];
                        if (string.IsNullOrEmpty(lineText))
                        {
                            // سطر فارغ
                            processedLines.Add(new LineInfo());
                            continue;
                        }

                        using (var buffer = new HarfBuzzSharp.Buffer())
                        {
                            buffer.Direction = Direction.RightToLeft;
                            buffer.Script = Script.Arabic;
                            buffer.Language = new Language("ar");
                            buffer.AddUtf8(lineText);
                            buffer.GuessSegmentProperties();

                            var features = new List<Feature>();
                            if (activeFeatures != null)
                            {
                                foreach (var f in activeFeatures)
                                    features.Add(new Feature(Tag.Parse(f), 1, 0, uint.MaxValue));
                            }

                            font.Shape(buffer, features.ToArray());

                            var glyphInfos = buffer.GlyphInfos;
                            var glyphPositions = buffer.GlyphPositions;

                            var lineInfo = new LineInfo();
                            float x = 0, y = 0;

                            for (int i = 0; i < glyphInfos.Length; i++)
                            {
                                ushort gid = (ushort)glyphInfos[i].Codepoint;
                                float px = x + (glyphPositions[i].XOffset * scale);
                                float py = y - (glyphPositions[i].YOffset * scale);

                                SKPath glyphPath = null;
                                try { glyphPath = skFont.GetGlyphPath(gid); } catch { glyphPath = null; }

                                if (glyphPath != null)
                                {
                                    var b = glyphPath.TightBounds;
                                    lineInfo.Glyphs.Add(new GlyphInfo
                                    {
                                        Codepoint = gid,
                                        X = px,
                                        Y = py,
                                        Path = glyphPath,
                                        TightBounds = b
                                    });

                                    float gminX = px + b.Left;
                                    float gmaxX = px + b.Right;
                                    float gminY = py + b.Top;
                                    float gmaxY = py + b.Bottom;

                                    if (gminX < lineInfo.MinX) lineInfo.MinX = gminX;
                                    if (gmaxX > lineInfo.MaxX) lineInfo.MaxX = gmaxX;
                                    if (gminY < lineInfo.MinY) lineInfo.MinY = gminY;
                                    if (gmaxY > lineInfo.MaxY) lineInfo.MaxY = gmaxY;
                                }

                                x += glyphPositions[i].XAdvance * scale;
                                y += glyphPositions[i].YAdvance * scale;
                            }

                            // إذا لم يحتوي السطر على حروف مرئية، نضع قيم افتراضية لتفادي أخطاء الأبعاد
                            if (lineInfo.MinX == float.MaxValue)
                            {
                                lineInfo.MinX = 0;
                                lineInfo.MaxX = 0;
                                lineInfo.MinY = 0;
                                lineInfo.MaxY = 0;
                            }

                            processedLines.Add(lineInfo);
                        }
                    }

                    // حساب العرض الأقصى للأسطر
                    float maxWidth = 0f;
                    foreach (var line in processedLines)
                    {
                        if (line.Width > maxWidth)
                            maxWidth = line.Width;
                    }

                    if (maxWidth <= 0)
                        maxWidth = 100f;

                    // تطبيق محاذاة النص وإزاحته رأسياً
                    float overallMinY = float.MaxValue;
                    float overallMaxY = float.MinValue;

                    var finalGlyphs = new List<(SKPath path, float finalX, float finalY)>();

                    for (int j = 0; j < processedLines.Count; j++)
                    {
                        var line = processedLines[j];
                        if (line.Glyphs.Count == 0) continue;

                        float alignX = 0f;
                        switch (alignment)
                        {
                            case TypoAlignment.Left:
                                alignX = -line.MinX;
                                break;
                            case TypoAlignment.Center:
                                alignX = -line.MinX + (maxWidth - line.Width) / 2f;
                                break;
                            case TypoAlignment.Right:
                            default:
                                alignX = -line.MinX + (maxWidth - line.Width);
                                break;
                        }

                        float alignY = j * lineHeight;

                        foreach (var glyph in line.Glyphs)
                        {
                            float fx = glyph.X + alignX;
                            float fy = glyph.Y + alignY;

                            finalGlyphs.Add((glyph.Path, fx, fy));

                            float gminY = fy + glyph.TightBounds.Top;
                            float gmaxY = fy + glyph.TightBounds.Bottom;

                            if (gminY < overallMinY) overallMinY = gminY;
                            if (gmaxY > overallMaxY) overallMaxY = gmaxY;
                        }
                    }

                    if (finalGlyphs.Count == 0)
                    {
                        File.WriteAllText(svgPath, "<svg xmlns='http://www.w3.org/2000/svg'></svg>");
                        return svgPath;
                    }

                    float padding = 10;
                    float svgW = maxWidth + padding * 2;
                    float svgH = (overallMaxY - overallMinY) + padding * 2;

                    using (var sw = new StreamWriter(svgPath, false, Encoding.UTF8))
                    {
                        sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                        sw.WriteLine(string.Format("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{0:F1}\" height=\"{1:F1}\" viewBox=\"0 0 {0:F1} {1:F1}\">", svgW, svgH));
                        sw.WriteLine("<g fill=\"black\" stroke=\"none\">");

                        foreach (var (path, fx, fy) in finalGlyphs)
                        {
                            string d = ConvertSkPathToSvgPath(path, fx + padding, fy - overallMinY + padding);
                            if (!string.IsNullOrWhiteSpace(d))
                                sw.WriteLine(string.Format("  <path d=\"{0}\"/>\n", d));
                        }

                        sw.WriteLine("</g>");
                        sw.WriteLine("</svg>");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CreateVectorText error: " + ex.Message);
            }

            return svgPath;
        }

        private static string ConvertSkPathToSvgPath(SKPath path, float offsetX, float offsetY)
        {
            if (path == null) return string.Empty;
            var sb = new StringBuilder();
            using (var it = path.CreateRawIterator())
            {
                SKPoint[] pts = new SKPoint[4];
                SKPathVerb v;
                while ((v = it.Next(pts)) != SKPathVerb.Done)
                {
                    switch (v)
                    {
                        case SKPathVerb.Move:
                            sb.AppendFormat("M{0:F2} {1:F2} ", pts[0].X + offsetX, pts[0].Y + offsetY);
                            break;
                        case SKPathVerb.Line:
                            sb.AppendFormat("L{0:F2} {1:F2} ", pts[1].X + offsetX, pts[1].Y + offsetY);
                            break;
                        case SKPathVerb.Quad:
                            sb.AppendFormat("Q{0:F2} {1:F2} {2:F2} {3:F2} ", pts[1].X + offsetX, pts[1].Y + offsetY, pts[2].X + offsetX, pts[2].Y + offsetY);
                            break;
                        case SKPathVerb.Cubic:
                            sb.AppendFormat("C{0:F2} {1:F2} {2:F2} {3:F2} {4:F2} {5:F2} ", pts[1].X + offsetX, pts[1].Y + offsetY, pts[2].X + offsetX, pts[2].Y + offsetY, pts[3].X + offsetX, pts[3].Y + offsetY);
                            break;
                        case SKPathVerb.Close:
                            sb.Append("Z ");
                            break;
                    }
                }
            }
            return sb.ToString();
        }

        public static BitmapSource RenderPreviewImage(string text, string fontPath, List<string> activeFeatures, SKColor textColor, TypoAlignment alignment)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
                return null;

            try
            {
                // فتح الملف كجداول بيانات ثنائية (Stream) لتفادي مشاكل الحروف العربية (Unicode) في مسارات الملفات بنظام ويندوز
                using (var fs1 = File.OpenRead(fontPath))
                using (var fs2 = File.OpenRead(fontPath))
                using (var typeface = SKTypeface.FromStream(fs1))
                using (var blob = Blob.FromStream(fs2))
                using (var face = new Face(blob, 0))
                using (var font = new HarfBuzzSharp.Font(face))
                {
                    // تفعيل دوال محرك OpenType وتحديد مقاييس الخط ليقوم HarfBuzz بتشكيل الحروف بشكل صحيح
                    font.SetFunctionsOpenType();
                    font.SetScale(face.UnitsPerEm, face.UnitsPerEm);

                    float fontSize = 80f;
                    var skFont = new SKFont(typeface, fontSize);
                    var paint = new SKPaint { Color = textColor, IsAntialias = true };
                    float scale = fontSize / face.UnitsPerEm;
                    var metrics = skFont.Metrics;

                    // تباعد الأسطر الرأسي (1.3x من ارتفاع الخط)
                    float lineHeight = (metrics.Descent - metrics.Ascent) * 1.3f;

                    string[] textLines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    var processedLines = new List<LineInfo>();

                    for (int j = 0; j < textLines.Length; j++)
                    {
                        string lineText = textLines[j];
                        if (string.IsNullOrEmpty(lineText))
                        {
                            processedLines.Add(new LineInfo());
                            continue;
                        }

                        using (var buffer = new HarfBuzzSharp.Buffer())
                        {
                            buffer.Direction = Direction.RightToLeft;
                            buffer.Script = Script.Arabic;
                            buffer.Language = new Language("ar");
                            buffer.AddUtf8(lineText);
                            buffer.GuessSegmentProperties();

                            var features = new List<Feature>();
                            if (activeFeatures != null)
                            {
                                foreach (var f in activeFeatures)
                                    features.Add(new Feature(Tag.Parse(f), 1, 0, uint.MaxValue));
                            }

                            font.Shape(buffer, features.ToArray());

                            var glyphInfos = buffer.GlyphInfos;
                            var glyphPositions = buffer.GlyphPositions;

                            var lineInfo = new LineInfo();
                            float x = 0, y = 0;

                            for (int i = 0; i < glyphInfos.Length; i++)
                            {
                                ushort gid = (ushort)glyphInfos[i].Codepoint;
                                float px = x + (glyphPositions[i].XOffset * scale);
                                float py = y - (glyphPositions[i].YOffset * scale);

                                lineInfo.Glyphs.Add(new GlyphInfo
                                {
                                    Codepoint = gid,
                                    X = px,
                                    Y = py,
                                    TightBounds = new SKRect() // سنحسب الأبعاد الكلية للرسم
                                });

                                // حساب أبعاد افتراضية تقديرية من خلال عرض الحرف
                                float gWidth = glyphPositions[i].XAdvance * scale;
                                float gminX = px;
                                float gmaxX = px + gWidth;

                                if (gminX < lineInfo.MinX) lineInfo.MinX = gminX;
                                if (gmaxX > lineInfo.MaxX) lineInfo.MaxX = gmaxX;

                                x += glyphPositions[i].XAdvance * scale;
                                y += glyphPositions[i].YAdvance * scale;
                            }

                            if (lineInfo.MinX == float.MaxValue)
                            {
                                lineInfo.MinX = 0;
                                lineInfo.MaxX = 0;
                            }

                            processedLines.Add(lineInfo);
                        }
                    }

                    // حساب العرض الأقصى للأسطر للمعاينة
                    float maxWidth = 0f;
                    foreach (var line in processedLines)
                    {
                        if (line.Width > maxWidth)
                            maxWidth = line.Width;
                    }

                    if (maxWidth <= 0)
                        maxWidth = 100f;

                    // بناء الـ TextBlob متعدد الأسطر مع محاذاته
                    var builder = new SKTextBlobBuilder();
                    
                    float overallMinY = float.MaxValue;
                    float overallMaxY = float.MinValue;

                    for (int j = 0; j < processedLines.Count; j++)
                    {
                        var line = processedLines[j];
                        if (line.Glyphs.Count == 0) continue;

                        var run = builder.AllocatePositionedRun(skFont, line.Glyphs.Count);
                        Span<ushort> runGlyphs = run.Glyphs;
                        Span<SKPoint> runPositions = run.Positions;

                        float alignX = 0f;
                        switch (alignment)
                        {
                            case TypoAlignment.Left:
                                alignX = -line.MinX;
                                break;
                            case TypoAlignment.Center:
                                alignX = -line.MinX + (maxWidth - line.Width) / 2f;
                                break;
                            case TypoAlignment.Right:
                            default:
                                alignX = -line.MinX + (maxWidth - line.Width);
                                break;
                        }

                        float alignY = j * lineHeight;

                        for (int i = 0; i < line.Glyphs.Count; i++)
                        {
                            var glyph = line.Glyphs[i];
                            runGlyphs[i] = glyph.Codepoint;
                            
                            float fx = glyph.X + alignX;
                            float fy = glyph.Y + alignY;
                            runPositions[i] = new SKPoint(fx, fy);

                            float gminY = fy + metrics.Ascent;
                            float gmaxY = fy + metrics.Descent;

                            if (gminY < overallMinY) overallMinY = gminY;
                            if (gmaxY > overallMaxY) overallMaxY = gmaxY;
                        }
                    }

                    using (var textBlob = builder.Build())
                    {
                        if (textBlob == null) return null;

                        float textWidth = maxWidth;
                        float textHeight = overallMaxY - overallMinY;

                        if (textWidth <= 0) textWidth = 100;
                        if (textHeight <= 5) textHeight = 50;

                        int padding = 20;
                        int width = (int)Math.Ceiling(textWidth) + padding * 2;
                        int height = (int)Math.Ceiling(textHeight) + padding * 2;

                        using (var bitmap = new SKBitmap(width, height))
                        {
                            using (var canvas = new SKCanvas(bitmap))
                            {
                                canvas.Clear(SKColors.Transparent);
                                
                                float drawX = padding;
                                float drawY = -overallMinY + padding;
                                canvas.DrawText(textBlob, drawX, drawY, paint);
                            }

                            using (var image = SKImage.FromBitmap(bitmap))
                            using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
                            using (var ms = new MemoryStream())
                            {
                                data.SaveTo(ms);
                                ms.Position = 0;

                                var bitmapImage = new BitmapImage();
                                bitmapImage.BeginInit();
                                bitmapImage.StreamSource = ms;
                                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                                bitmapImage.EndInit();
                                bitmapImage.Freeze();

                                return bitmapImage;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error rendering preview image: " + ex.Message);
                return null;
            }
        }
    }
}