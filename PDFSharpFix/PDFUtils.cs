using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace PDFSharpFix
{
    ///
    /// A utility class responsible for packaging a list of Images into a PDF file using PDFSharp library.
    ///
    public class PDFUtils
    {
        /// <summary>
        /// Do simple packaging of list of images into a PDF
        /// </summary>
        ///
        /// <param name="pages">List of pages</param>
        /// <param name="subject">PDF subject</param>
        /// <param name="title">PDF title</param>
        /// <returns></returns>
        public static byte[] BuildPDFInPDFSharp(IList<System.Drawing.Image> pages, string subject, string title)
        {
            using (var ms = new MemoryStream())
            {
                PdfSharp.Pdf.PdfDocument doc = new PdfSharp.Pdf.PdfDocument(ms);
                int i = 1;
                
                foreach (var image in pages)
                {
                    PdfSharp.Pdf.PdfPage page = doc.AddPage();

                    // Set page size
                    //page.Size = PdfSharp.PageSize.Undefined;
                    page.Width = Math.Round(image.PhysicalDimension.Width / image.HorizontalResolution * 72, 0);
                    page.Height = Math.Round(image.PhysicalDimension.Height / image.HorizontalResolution * 72, 0);

                    // Render image to the page
                    PdfSharp.Drawing.XGraphics gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page);

                    var ximage = PdfSharp.Drawing.XImage.FromGdiPlusImage(PDFUtils.ConvertImageToTiff(image));
                    gfx.DrawImage(image, new PdfSharp.Drawing.XPoint(0, 0));
                    i++;
                }

                doc.Info.Author = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                doc.Info.Creator = "PDFSharpFix";
                doc.Info.Subject = subject;
                doc.Info.Title = title;

                doc.Save(ms);
                ms.Flush();

                var buffer = ms.GetBuffer();

                return buffer;
            }
        }

        protected static System.Drawing.Image ConvertImageToTiff(System.Drawing.Image image)
        {
            System.Drawing.Image result;

            using(var memStream = new MemoryStream())
            {
                lock (image)
                {
                    image.Save(memStream, ImageFormat.Tiff);
                }

                memStream.Flush();

                memStream.Seek(0, SeekOrigin.Begin);

                result = System.Drawing.Image.FromStream(memStream);

                return result;
            }
        }
    }
}
