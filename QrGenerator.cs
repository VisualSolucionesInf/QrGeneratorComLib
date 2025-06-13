using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using QRCoder;

namespace QrGeneratorComLib
{
    [ComVisible(true)]
    [Guid("B1111111-2222-3333-4444-555555555555")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IQrGenerator
    {
        byte[] GenerateQrBytes(string text, string extension = "PNG");
        string GenerateQrBase64(string text, string extension = "PNG");
        void SaveQrToFile(string text, string filePath, int width, int height, string extension = "PNG");
    }

    [ComVisible(true)]
    [Guid("A1111111-2222-3333-4444-555555555555")]
    [ClassInterface(ClassInterfaceType.None)]
    public class QrGenerator : IQrGenerator
    {
        private Bitmap GenerateQrBitmap(string text, int? width = null, int? height = null)
        {
            var qrGenerator = new QRCodeGenerator();
            var data = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q);
            var qrCode = new QRCode(data);
            Bitmap bmp = qrCode.GetGraphic(20); // scale factor 20

            if (width.HasValue && height.HasValue)
            {
                return new Bitmap(bmp, new Size(width.Value, height.Value));
            }

            return bmp;
        }

        public byte[] GenerateQrBytes(string text, string extension = "PNG")
        {
            using (var bmp = GenerateQrBitmap(text))
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, GetImageFormat(extension));
                return ms.ToArray();
            }
        }

        public string GenerateQrBase64(string text, string extension = "PNG")
        {
            using (var bmp = GenerateQrBitmap(text))
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, GetImageFormat(extension));
                return Convert.ToBase64String(ms.ToArray());
            }
        }

        public void SaveQrToFile(string text, string filePath, int width, int height, string extension = "PNG")
        {
            using (var bmp = GenerateQrBitmap(text, width, height))
            {
                bmp.Save(filePath, GetImageFormat(extension));
            }
        }

        protected ImageFormat GetImageFormat(string extension)
        {
            switch (extension.Trim().ToUpper())
            {
                case "PNG":
                    return ImageFormat.Png;
                case "BMP":
                    return ImageFormat.Bmp;
                case "JPG":
                case "JPEG":
                    return ImageFormat.Jpeg;
                default:
                    return ImageFormat.Png;
            }
        }
    }
}
