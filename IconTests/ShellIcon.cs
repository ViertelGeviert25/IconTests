using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;

namespace IconTests
{
    /// <summary>
    /// Provides method for retrieving the System.Drawing.Icon from a file
    /// </summary>
    public static class ShellIcon
    {
        // https://stackoverflow.com/questions/8499633/how-to-display-base64-images-in-html
        // https://stackoverflow.com/questions/10889764/how-to-convert-bitmap-to-a-base64-string
        // https://www.pinvoke.net/default.aspx/shell32.SHGetFileInfo
        // https://stackoverflow.com/questions/28525925/get-icon-128128-file-type-c-sharp?noredirect=1&lq=1


        private const string IID_IImageList = "46EB5926-582E-4017-9FDF-E8998DAA0950";
        private const string DEFAULT_WINDOWS_FONT = "Segoe UI";

        [Flags]
        private enum SHGFI : uint
        {
            /// <summary>get icon</summary>
            Icon = 0x000000100,
            /// <summary>get display name</summary>
            DisplayName = 0x000000200,
            /// <summary>get type name</summary>
            TypeName = 0x000000400,
            /// <summary>get attributes</summary>
            Attributes = 0x000000800,
            /// <summary>get icon location</summary>
            IconLocation = 0x000001000,
            /// <summary>return exe type</summary>
            ExeType = 0x000002000,
            /// <summary>get system icon index</summary>
            SysIconIndex = 0x000004000,
            /// <summary>put a link overlay on icon</summary>
            LinkOverlay = 0x000008000,
            /// <summary>show icon in selected state</summary>
            Selected = 0x000010000,
            /// <summary>get only specified attributes</summary>
            Attr_Specified = 0x000020000,
            /// <summary>get large icon</summary>
            LargeIcon = 0x000000000,
            /// <summary>get small icon</summary>
            SmallIcon = 0x000000001,
            /// <summary>get open icon</summary>
            OpenIcon = 0x000000002,
            /// <summary>get shell size icon</summary>
            ShellIconSize = 0x000000004,
            /// <summary>pszPath is a pidl</summary>
            PIDL = 0x000000008,
            /// <summary>use passed dwFileAttribute</summary>
            UseFileAttributes = 0x000000010,
            /// <summary>apply the appropriate overlays</summary>
            AddOverlays = 0x000000020,
            /// <summary>Get the index of the overlay in the upper 8 bits of the iIcon</summary>
            OverlayIndex = 0x000000040,
        }

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


        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left, top, right, bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            int x;
            int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGELISTDRAWPARAMS
        {
            public int cbSize;
            public IntPtr himl;
            public int i;
            public IntPtr hdcDst;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public int xBitmap;    // x offest from the upperleft of bitmap
            public int yBitmap;    // y offset from the upperleft of bitmap
            public int rgbBk;
            public int rgbFg;
            public int fStyle;
            public int dwRop;
            public int fState;
            public int Frame;
            public int crEffect;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct IMAGEINFO
        {
            public IntPtr hbmImage;
            public IntPtr hbmMask;
            public int Unused1;
            public int Unused2;
            public RECT rcImage;
        }
        [ComImportAttribute()]
        [GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950")]
        [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IImageList
        {
            [PreserveSig]
            int Add(
            IntPtr hbmImage,
            IntPtr hbmMask,
            ref int pi);

            [PreserveSig]
            int ReplaceIcon(
            int i,
            IntPtr hicon,
            ref int pi);

            [PreserveSig]
            int SetOverlayImage(
            int iImage,
            int iOverlay);

            [PreserveSig]
            int Replace(
            int i,
            IntPtr hbmImage,
            IntPtr hbmMask);

            [PreserveSig]
            int AddMasked(
            IntPtr hbmImage,
            int crMask,
            ref int pi);

            [PreserveSig]
            int Draw(
            ref IMAGELISTDRAWPARAMS pimldp);

            [PreserveSig]
            int Remove(
            int i);

            [PreserveSig]
            int GetIcon(
            int i,
            int flags,
            ref IntPtr picon);

            [PreserveSig]
            int GetImageInfo(
            int i,
            ref IMAGEINFO pImageInfo);

            [PreserveSig]
            int Copy(
            int iDst,
            IImageList punkSrc,
            int iSrc,
            int uFlags);

            [PreserveSig]
            int Merge(
            int i1,
            IImageList punk2,
            int i2,
            int dx,
            int dy,
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int Clone(
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int GetImageRect(
            int i,
            ref RECT prc);

            [PreserveSig]
            int GetIconSize(
            ref int cx,
            ref int cy);

            [PreserveSig]
            int SetIconSize(
            int cx,
            int cy);

            [PreserveSig]
            int GetImageCount(
            ref int pi);

            [PreserveSig]
            int SetImageCount(
            int uNewCount);

            [PreserveSig]
            int SetBkColor(
            int clrBk,
            ref int pclr);

            [PreserveSig]
            int GetBkColor(
            ref int pclr);

            [PreserveSig]
            int BeginDrag(
            int iTrack,
            int dxHotspot,
            int dyHotspot);

            [PreserveSig]
            int EndDrag();

            [PreserveSig]
            int DragEnter(
            IntPtr hwndLock,
            int x,
            int y);

            [PreserveSig]
            int DragLeave(
            IntPtr hwndLock);

            [PreserveSig]
            int DragMove(
            int x,
            int y);

            [PreserveSig]
            int SetDragCursorImage(
            ref IImageList punk,
            int iDrag,
            int dxHotspot,
            int dyHotspot);

            [PreserveSig]
            int DragShowNolock(
            int fShow);

            [PreserveSig]
            int GetDragImage(
            ref POINT ppt,
            ref POINT pptHotspot,
            ref Guid riid,
            ref IntPtr ppv);

            [PreserveSig]
            int GetItemFlags(
            int i,
            ref int dwFlags);

            [PreserveSig]
            int GetOverlayImage(
            int iOverlay,
            ref int piIndex);
        };

        private class Shell32
        {
            public const uint SHGFI_ICON = 0x100;
            public const uint SHGFI_LARGEICON = 0x0; // 'Large icon
            public const uint SHGFI_SMALLICON = 0x1; // 'Small icon

            public const int SHIL_LARGE = 0x0;
            public const int SHIL_SMALL = 0x1;
            public const int SHIL_EXTRALARGE = 0x2;
            public const int SHIL_SYSSMALL = 0x3;
            public const int SHIL_JUMBO = 0x4;
            public const int SHIL_LAST = 0x4;

            public const int ILD_TRANSPARENT = 0x00000001;
            public const int ILD_IMAGE = 0x00000020;

            [DllImport("shell32.dll", EntryPoint = "#727")]
            public extern static int SHGetImageList(int iImageList, ref Guid riid, ref IImageList ppv);

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

        public static string GetFileIconAsBase64(string filePath)
        {
            IImageList spiml = null;
            Guid guil = new Guid(IID_IImageList);//or IID_IImageList
            Shell32.SHGetImageList(Shell32.SHIL_EXTRALARGE, ref guil, ref spiml);
            IntPtr hIcon = IntPtr.Zero;
            SHFILEINFO sfi = new SHFILEINFO();
            var sizeFileInfo = (uint)Marshal.SizeOf(sfi);
            var flags = (uint)(SHGFI.SysIconIndex | SHGFI.LargeIcon | SHGFI.UseFileAttributes);
            Shell32.SHGetFileInfo(filePath, 0, ref sfi, sizeFileInfo, flags);
            var iImage = sfi.iIcon;
            spiml.GetIcon(iImage, Shell32.ILD_TRANSPARENT | Shell32.ILD_IMAGE, ref hIcon);


            //SHFILEINFO shinfo = new SHFILEINFO();
            //var sizeFileInfo = (uint)Marshal.SizeOf(shinfo);
            //var flags = Shell32.SHGFI_ICON | Shell32.SHGFI_LARGEICON;
            //Shell32.SHGetFileInfo(filePath, 0, ref shinfo, sizeFileInfo, flags);
            //var hIcon = shinfo.iIcon;

            // Convert the symbol to a bitmap
            Icon icon = Icon.FromHandle(hIcon);
            Bitmap iconBitmap = icon.ToBitmap();

            // Create a new bitmap containing the symbol and the file name
            int width = Math.Max(iconBitmap.Width, 200); // Width of the image
            // Height of the image (symbol height + space for the file name)
            int height = iconBitmap.Height + 30;

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
                    string fileName = Path.GetFileName(filePath);
                    Font font = new Font(DEFAULT_WINDOWS_FONT, 10);
                    SizeF textSize = g.MeasureString(fileName, font);

                    // Position of the text (centred under the symbol)
                    float textX = (width - textSize.Width) / 2;
                    float textY = iconBitmap.Height + 5;

                    g.DrawString(fileName, font, Brushes.Black, new PointF(textX, textY));
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
