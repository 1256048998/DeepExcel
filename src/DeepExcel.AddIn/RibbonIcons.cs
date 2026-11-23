using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace DeepExcel.AddIn
{
    /// <summary>
    /// 功能区按钮图标，运行时按屏幕 DPI 绘制。
    ///
    /// 以前用 imageMso（WindowUnhide / Info）：Excel 里是通用系统图标，跟品牌无关；
    /// WPS 表格没有 WindowUnhide，直接显示成一个问号方块（2026-11-23 真机）。
    /// 现在两个宿主用同一套自绘图标：
    ///   打开面板 = DeepExcel 标志（与面板左上角 Logo.tsx 的 LogoMark 同一图形）
    ///   帮助     = 同色系圆形里一个白色问号
    /// 不放 PNG 资源：矢量绘制在 125% / 150% / 200% 缩放下都清晰，也不用改资源管线。
    /// </summary>
    internal static class RibbonIcons
    {
        private static readonly Color BrandLight = Color.FromArgb(0x3b, 0x82, 0xf6);
        private static readonly Color BrandDark = Color.FromArgb(0x1a, 0x55, 0xc4);

        /// <summary>按按钮 id 返回 IPictureDisp（Office / WPS 的 getImage 回调要这个类型）。</summary>
        public static object ForControl(string controlId)
        {
            int size = LargeIconPixels();
            using (var bmp = controlId == "btnHelp" ? DrawHelp(size) : DrawLogo(size))
            {
                return PictureConverter.ToPictureDisp(bmp);
            }
        }

        /// <summary>大按钮 32 逻辑像素，乘以系统 DPI 缩放。</summary>
        internal static int LargeIconPixels()
        {
            float dpi = 96f;
            try
            {
                using (var g = Graphics.FromHwnd(IntPtr.Zero)) dpi = g.DpiX;
            }
            catch { }
            int px = (int)Math.Round(32 * dpi / 96f);
            return Math.Max(32, Math.Min(128, px));
        }

        /// <summary>DeepExcel 标志，24×24 设计稿按比例放大。</summary>
        internal static Bitmap DrawLogo(int size)
        {
            var bmp = NewCanvas(size);
            using (var g = Graphics.FromImage(bmp))
            {
                Prepare(g);
                float s = size / 24f;
                // 留 1/24 的边，避免贴边被功能区裁掉
                g.TranslateTransform(0.5f * s, 0.5f * s);
                g.ScaleTransform(23f / 24f, 23f / 24f);

                using (var tile = RoundedRect(0, 0, 24 * s, 24 * s, 6.5f * s))
                using (var fill = new LinearGradientBrush(new PointF(0, 0), new PointF(24 * s, 24 * s), BrandLight, BrandDark))
                {
                    g.FillPath(fill, tile);
                }

                using (var solid = new SolidBrush(Color.White))
                using (var soft = new SolidBrush(Color.FromArgb(158, 255, 255, 255)))
                {
                    FillRounded(g, solid, 5 * s, 5 * s, 6 * s, 6 * s, 1.6f * s);
                    FillRounded(g, soft, 13 * s, 5 * s, 6 * s, 6 * s, 1.6f * s);
                    FillRounded(g, soft, 5 * s, 13 * s, 6 * s, 6 * s, 1.6f * s);

                    // AI 星芒：Logo.tsx 里那条四段三次贝塞尔路径
                    using (var spark = new GraphicsPath())
                    {
                        spark.AddBezier(P(16, 12.2f, s), P(16.35f, 14.3f, s), P(17.2f, 15.25f, s), P(19.8f, 16, s));
                        spark.AddBezier(P(19.8f, 16, s), P(17.2f, 16.75f, s), P(16.35f, 17.7f, s), P(16, 19.8f, s));
                        spark.AddBezier(P(16, 19.8f, s), P(15.65f, 17.7f, s), P(14.8f, 16.75f, s), P(12.2f, 16, s));
                        spark.AddBezier(P(12.2f, 16, s), P(14.8f, 15.25f, s), P(15.65f, 14.3f, s), P(16, 12.2f, s));
                        spark.CloseFigure();
                        g.FillPath(solid, spark);
                    }
                }
            }
            return bmp;
        }

        /// <summary>帮助：品牌蓝圆形 + 白色问号。</summary>
        internal static Bitmap DrawHelp(int size)
        {
            var bmp = NewCanvas(size);
            using (var g = Graphics.FromImage(bmp))
            {
                Prepare(g);
                float s = size / 24f;
                var circle = new RectangleF(1.5f * s, 1.5f * s, 21 * s, 21 * s);
                using (var fill = new LinearGradientBrush(new PointF(0, 0), new PointF(24 * s, 24 * s), BrandLight, BrandDark))
                {
                    g.FillEllipse(fill, circle);
                }

                // 问号：弧 + 竖笔画 + 圆点，线宽与标志的格子圆角一个量级
                using (var pen = new Pen(Color.White, 2.4f * s) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                using (var path = new GraphicsPath())
                {
                    // 上半圆弧：圆心 (12, 9.6)，半径 3.3，从左侧 200° 顺时针扫到右下 40°
                    path.AddArc(new RectangleF(8.7f * s, 6.3f * s, 6.6f * s, 6.6f * s), 195f, 225f);
                    path.AddBezier(P(14.53f, 11.72f, s), P(13.4f, 12.7f, s), P(12f, 13.2f, s), P(12f, 14.9f, s));
                    g.DrawPath(pen, path);
                }
                using (var dot = new SolidBrush(Color.White))
                {
                    float r = 1.45f * s;
                    g.FillEllipse(dot, 12 * s - r, 18.1f * s - r, 2 * r, 2 * r);
                }
            }
            return bmp;
        }

        private static Bitmap NewCanvas(int size)
        {
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            bmp.MakeTransparent();
            return bmp;
        }

        private static void Prepare(Graphics g)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            g.Clear(Color.Transparent);
        }

        private static PointF P(float x, float y, float s) => new PointF(x * s, y * s);

        private static void FillRounded(Graphics g, Brush brush, float x, float y, float w, float h, float r)
        {
            using (var path = RoundedRect(x, y, w, h, r)) g.FillPath(brush, path);
        }

        private static GraphicsPath RoundedRect(float x, float y, float w, float h, float r)
        {
            float d = r * 2;
            var path = new GraphicsPath();
            path.AddArc(x, y, d, d, 180, 90);
            path.AddArc(x + w - d, y, d, d, 270, 90);
            path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
            path.AddArc(x, y + h - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        /// <summary>AxHost 的 protected 静态方法是 .NET 里把 Image 转成 IPictureDisp 的标准途径。</summary>
        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base(string.Empty) { }

            public static object ToPictureDisp(Image image) => GetIPictureDispFromPicture(image);
        }
    }
}
