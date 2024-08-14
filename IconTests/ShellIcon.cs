using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace IconTests
{
    /// <summary>
    /// Provides method for retrieving the System.Drawing.Icon from a file
    /// </summary>
    public static class ShellIcon
    {
        private const string DEFAULT_WINDOWS_FONT = "Segoe UI";
        private const int TEXT_MAX_WIDTH_PER_LINE = 200;
        private const int TEXT_MIN_WIDTH_PER_LINE = 100;

        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO
        {
            public const int NAMESIZE = 80;
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };

        private class Shell32
        {
            public const uint SHGFI_ICON = 0x100;
            public const uint SHGFI_LARGEICON = 0x0; // 'Large icon
            public const uint SHGFI_SMALLICON = 0x1; // 'Small icon

            //[DllImport("shell32.dll", EntryPoint = "#727")]
            //public extern static int SHGetImageList(int iImageList, ref Guid riid, ref IImageList ppv);

            [DllImport("user32.dll", EntryPoint = "DestroyIcon", SetLastError = true)]
            public static extern int DestroyIcon(IntPtr hIcon);

            [DllImport("Shell32.dll")]
            public static extern IntPtr SHGetFileInfo(
                string pszPath,
                uint dwFileAttributes,
                ref SHFILEINFO psfi,
                uint cbFileInfo,
                uint uFlags
            );
        }

        private static SizeF GetTextSize(string text, Font font)
        {
            Image fakeImage = new Bitmap(1, 1);
            Graphics graphics = Graphics.FromImage(fakeImage);
            SizeF size = graphics.MeasureString(text, font);
            return size;
        }


        public static int DrawWrappedText(Graphics graphics, string text, Font font, Brush brush, RectangleF layoutRect, ref PointF point, bool draw)
        {
            StringFormat stringFormat = new StringFormat
            {
                LineAlignment = StringAlignment.Near,
                Alignment = StringAlignment.Center,
                FormatFlags = StringFormatFlags.NoWrap
            };

            float yOffset = point.Y;

            // Split the text into lines that fit within the layoutRect width
            string[] words = text.Split(' ');
            string line = string.Empty;

            var lineCounter = 1;

            var subword = string.Empty;
            var sep = string.Empty;
            foreach (string word in words)
            {
                subword += word + " ";
                string testLine = line + (line.Length > 0 ? sep : "") + word;
                SizeF textSize = graphics.MeasureString(testLine, font);

                if (textSize.Width > layoutRect.Width)
                {
                    if (!string.IsNullOrEmpty(line))
                    {
                        line += sep;
                        // Draw the current line
                        graphics.DrawString(line, font, brush, new PointF(point.X, yOffset), stringFormat);
                        yOffset += textSize.Height; // Move down for the next line
                        lineCounter++;
                    }

                    // Start a new line with the current word
                    line = word;
                }
                else
                {
                    // Add the word to the current line
                    line = testLine;
                }
                sep = subword.Length < text.Length ? text.Substring(subword.Length - 1, 1) : string.Empty;
            }

            // Draw the last line
            if (line.Length > 0)
            {
                graphics.DrawString(line, font, brush, new PointF(point.X, yOffset), stringFormat);
            }

            // Update the point to the new position after drawing
            point.Y = yOffset + font.GetHeight(graphics); // Adjust for some margin if needed

            return lineCounter;
        }

        public static string GetFileIconAsBase64(string filePath, string displayName)
        {
            SHFILEINFO shinfo = new SHFILEINFO();
            var sizeFileInfo = (uint)Marshal.SizeOf(shinfo);
            var flags = Shell32.SHGFI_ICON | Shell32.SHGFI_LARGEICON;
            Shell32.SHGetFileInfo(filePath, 0, ref shinfo, sizeFileInfo, flags);
            var hIcon = shinfo.hIcon;

            // Convert the symbol to a bitmap
            Icon icon = Icon.FromHandle(hIcon);
            Bitmap iconBitmap = icon.ToBitmap();

            Font font = new Font(DEFAULT_WINDOWS_FONT, 9);
            SizeF textSize = GetTextSize(displayName, font);

            // Create a new bitmap containing the symbol and the file name
            int width = Math.Max(iconBitmap.Width, (int)textSize.Width); // Width of the image
            if (width > TEXT_MAX_WIDTH_PER_LINE)
            {
                width = TEXT_MAX_WIDTH_PER_LINE;
            }

            float textX = (width) / 2;
            float textY = 0;

            PointF initialPoint = new PointF(Math.Abs(textX), textY);
            var lines = DrawWrappedText(Graphics.FromImage(new Bitmap(1, 1)),
                displayName, font, Brushes.Black, new RectangleF(0, 0, width, 1000), ref initialPoint, false);

            // Height of the image (symbol height + space for the file name)
            int height = iconBitmap.Height + (5 + (int)font.GetHeight()) * lines + 5;

            var sigBase64 = string.Empty;

            using (Bitmap resultBitmap = new Bitmap(width, height))
            {
                using (Graphics g = Graphics.FromImage(resultBitmap))
                {
                    // Fill background white
                    g.Clear(Color.White);

                    // Draw the symbol
                    g.DrawImage(iconBitmap, (width - iconBitmap.Width) / 2, 0);

                    // Draw the file name under the symbol

                    // Position of the text (centred under the symbol)
                    textX = (width) / 2;
                    textY = iconBitmap.Height + 5;

                    initialPoint = new PointF(Math.Abs(textX), textY);

                    //g.DrawString(displayName, font, Brushes.Black, new PointF(textX, textY));
                    DrawWrappedText(g,
                        displayName, font, Brushes.Black, new RectangleF(0, 0, width, height), ref initialPoint, true);
                }

                var ms = new MemoryStream();
                resultBitmap.Save(ms, ImageFormat.Png);
                var byteImage = ms.ToArray();
                sigBase64 = Convert.ToBase64String(byteImage); // Get Base64
            }

            // Release icon handle
            Shell32.DestroyIcon(hIcon);
            return sigBase64;
        }
    }
}
