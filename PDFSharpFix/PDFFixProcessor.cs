using BitMiracle.LibTiff.Classic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Filters;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace PDFSharpFix
{
    /// <summary>
    /// Utility class responsible for fixing bitmaps broken by old version of PDFSharp.
    /// </summary>
    public class PDFFixProcessor
    {
        static FlateDecode flateDecoder = new FlateDecode();

        /// <summary>
        /// Older versions of PDFSharp used to create flawed instances of CCITT G4 encoded B/W bitmaps.
        /// In order to fix these, they need to be processed by a reader that encodes them again.
        /// </summary>
        /// <param name="pdfStream"></param>
        /// <returns></returns>
        public static byte[] FixPDFImages(Stream pdfStream)
        {
            var pdfDocument = PdfReader.Open(pdfStream);

            int imageCount = 0;

            var images = new List<Image>();

            foreach (PdfPage page in pdfDocument.Pages)
            {
                // Get resources dictionary
                PdfDictionary resources = page.Elements.GetDictionary("/Resources");
                if (resources != null)
                {
                    // Get external objects dictionary
                    PdfDictionary xObjects = resources.Elements.GetDictionary("/XObject");
                    if (xObjects != null)
                    {
                        ICollection<PdfItem> items = xObjects.Elements.Values;
                        // Iterate references to external objects
                        foreach (PdfItem item in items)
                        {
                            PdfReference reference = item as PdfReference;
                            if (reference != null)
                            {
                                PdfDictionary xObject = reference.Value as PdfDictionary;
                                // Is external object an image?
                                if (xObject != null && xObject.Elements.GetString("/Subtype") == "/Image")
                                {
                                    imageCount++;
                                    images.Add(ExportImage(xObject, ref imageCount));
                                }
                            }
                        }
                    }
                }
            }

            byte[] pdfRaw = PDFUtils.BuildPDFInPDFSharp(images, pdfDocument.Info.Subject, pdfDocument.Info.Title);

            return pdfRaw;
        }

        protected static Image ExportImage(PdfDictionary image, ref int count)
        {
            if (image.Elements.ContainsKey("/Filter"))
            {
                string filter = "";

                try
                {
                    filter = image.Elements.GetName("/Filter");
                }
                catch (Exception exc) // swallow exception - this is because we are looking at an array
                {
                    var array = image.Elements.GetArray("/Filter");
                    for (int i=0; i< array.Count(); i++)
                    {
                        filter += array.Elements.GetName(i);
                    }
                }

                Image correctedPage = null;

                switch (filter)
                {
                    case "/DCTDecode":
                        correctedPage = ExportJpegImage(image, ref count);
                        break;
                    case "/FlateDecode":
                        correctedPage = ExportAsImageFlate(image, ref count);
                        break;
                    case "/FlateDecode/CCITTFaxDecode":
                        correctedPage = ExportAs1BitTiffImageCCITT(image, ref count, true);
                        break;
                    case "/CCITTFaxDecode":
                        correctedPage = ExportAs1BitTiffImageCCITT(image, ref count, false);
                        break;
                    default:
                        throw new Exception(String.Format("Unknown filter combination: [{0}]", filter));
                }

                return correctedPage;
            }

            return null;
        }

        static Bitmap ExportJpegImage(PdfDictionary image, ref int count)
        {
            // Fortunately JPEG has native support in PDF and exporting an image is just writing the stream to a file.
            byte[] stream = image.Stream.Value;
            using (var memStream = new MemoryStream(stream))
            {
                var result = new Bitmap(memStream);

                return result;
            }
        }

        static Bitmap ExportAsImageFlate(PdfDictionary image, ref int count)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);
            int xres = 200;
            int yres = 200;
            int bitsPerComponent = image.Elements.GetInteger(PdfImage.Keys.BitsPerComponent);
            string colorSpace = image.Elements.GetName(PdfImage.Keys.ColorSpace);
            byte[] stream = image.Stream.Value;

            // first we need to decode this stream, then put it in TIFF file
            var decodedStream = flateDecoder.Decode(stream);
            
            int rowsize = 0;
            PixelFormat format = PixelFormat.Undefined;

            if (bitsPerComponent == 1)
            {
                rowsize = width / 8 + 1;
                format = PixelFormat.Format1bppIndexed;
            }

            if (bitsPerComponent == 8)
            {
                if (colorSpace == "/DeviceRGB")
                {
                    rowsize = width * 3;
                    format = PixelFormat.Format24bppRgb;
                }
                else
                {
                    rowsize = width;
                    format = PixelFormat.Format8bppIndexed;
                }
            }

            // Create destination bitmap
            Bitmap destination = new Bitmap(width, height, format);
            destination.SetResolution(xres, yres);

            // Lock destination bitmap in memory
            BitmapData destinationData = destination.LockBits(new Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, format);

            int imageSize = destinationData.Stride * height;
            byte[] destinationBuffer = new byte[imageSize];
            int destinationIndex = 0;

            for (int i = 0; i < height; i++)
            {
                switch(format)
                {
                    case PixelFormat.Format24bppRgb:
                        destinationIndex = i * destinationData.Stride;
                        for (int x = 0; x < width; x++)
                        {
                            destinationBuffer[destinationIndex] = decodedStream[i * rowsize + x * 3 + 2];
                            destinationBuffer[destinationIndex+1] = decodedStream[i * rowsize + x * 3 + 1];
                            destinationBuffer[destinationIndex+2] = decodedStream[i * rowsize + x * 3];
                            destinationIndex+=3;
                        }
                        break;
                    case PixelFormat.Format1bppIndexed:
                        destinationIndex = i * destinationData.Stride;
                        Array.Copy(decodedStream, i * rowsize, destinationBuffer, i * destinationData.Stride, rowsize);
                        break;
                }
            }

            // Copy binary image data to destination bitmap
            Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, destinationBuffer.Length);

            // Unlock destination bitmap
            destination.UnlockBits(destinationData);

            return destination;
        }

        static Bitmap ExportAs1BitTiffImageCCITT(PdfDictionary image, ref int count, bool isStreamCompressed)
        {
            int width = image.Elements.GetInteger(PdfImage.Keys.Width);
            int height = image.Elements.GetInteger(PdfImage.Keys.Height);
            int bitsPerComponent = image.Elements.GetInteger(PdfImage.Keys.BitsPerComponent);
            int xres = 200;
            int yres = 200;
            byte[] stream = image.Stream.Value;

            // if stream compressed, first we need to decode this stream, then put it in TIFF file
            if (isStreamCompressed)
            {
                stream = flateDecoder.Decode(stream);
            }

            // generate temp file name
            string name = Path.GetTempFileName();

            // Stream the bits into a tiff file so thet they can be decoded later
            // I know this looks cheesy, but I don't have time to do it in-memory...
            var tiff = Tiff.Open(name, "w");
            tiff.SetField(TiffTag.IMAGEWIDTH, "" + width);
            tiff.SetField(TiffTag.IMAGELENGTH, "" + height);
            tiff.SetField(TiffTag.COMPRESSION, "" + (int)Compression.CCITTFAX4);
            tiff.SetField(TiffTag.BITSPERSAMPLE, "" + bitsPerComponent);
            tiff.SetField(TiffTag.SAMPLESPERPIXEL, 1);

            tiff.SetField(TiffTag.FAXMODE, "" + (int)FaxMode.CLASSF);

            tiff.WriteRawStrip(0, stream, stream.Length);
            tiff.Close();

            int i = 0, j = 0;

            // create destination bitmap
            Bitmap destination = new Bitmap(width, height, PixelFormat.Format1bppIndexed);
            destination.SetResolution(xres, yres);

            // Lock destination bitmap in memory
            BitmapData destinationData = destination.LockBits(new Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);

            var tiff2 = Tiff.Open(name, "r");

            int rowsize = tiff2.ScanlineSize();
            int imageSize = destinationData.Stride * height;
            byte[] destinationBuffer = new byte[imageSize];
            byte[] scanline = new byte[rowsize];

            int destinationIndex = 0;

            for (i = 0; i < height; i++)
            {
                destinationIndex = i * destinationData.Stride;

                var readResult = tiff2.ReadScanline(scanline, i);

                for (j = 0; j < rowsize; j++)
                {
                    destinationBuffer[destinationIndex] = (byte)~(scanline[j]); // "not" in order to get correct colors
                    destinationIndex++;
                }
            }

            tiff2.Close();

            if (File.Exists(name))
            {
                File.Delete(name);
            }

            // Copy binary image data to destination bitmap
            Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize);

            // Unlock destination bitmap
            destination.UnlockBits(destinationData);

            return destination;

        }

    }
}
