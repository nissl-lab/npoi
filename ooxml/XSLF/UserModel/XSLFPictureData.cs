using System;
using System.IO;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.Util;
using SkiaSharp;

namespace NPOI.XSLF.UserModel
{
    /// <summary>
    /// Read-only image bytes stored directly in a presentation package part.
    /// </summary>
    public class XSLFPictureData : POIXMLDocumentPart, IPictureData
    {
        internal XSLFPictureData()
        {
            Committed = true;
        }

        internal XSLFPictureData(POIXMLDocumentPart parent, PackagePart part) : base(parent, part)
        {
            Committed = true;
        }

        /// <summary>
        /// Gets a fresh copy of the original image bytes, without conversion.
        /// </summary>
        public byte[] Data
        {
            get
            {
                EnsureAttached();
                using var input = GetPackagePart().GetInputStream();
                return IOUtils.ToByteArray(input);
            }
        }

        /// <inheritdoc/>
        public string MimeType => GetPackagePart().ContentType;

        /// <inheritdoc/>
        public PictureType PictureType
        {
            get
            {
                switch (MimeType.ToLowerInvariant())
                {
                    case "image/png": return PictureType.PNG;
                    case "image/jpeg":
                    case "image/jpg": return PictureType.JPEG;
                    case "image/x-emf": return PictureType.EMF;
                    case "image/x-wmf": return PictureType.WMF;
                    case "image/pict": return PictureType.PICT;
                    case "image/dib": return PictureType.DIB;
                    case "image/gif": return PictureType.GIF;
                    case "image/tiff": return PictureType.TIFF;
                    case "image/x-eps": return PictureType.EPS;
                    case "image/bmp":
                    case "image/x-ms-bmp": return PictureType.BMP;
                    case "image/x-wpg": return PictureType.WPG;
                    default: return PictureType.None;
                }
            }
        }

        /// <inheritdoc/>
        public string SuggestFileExtension()
        {
            return GetPackagePart().PartName.Extension;
        }

        /// <summary>
        /// Returns the encoded PNG/JPEG pixel dimensions, without applying DPI or EXIF orientation.
        /// Invalid, incomplete, mismatched and unsupported images fail explicitly.
        /// </summary>
        public SKSizeI GetImageDimensionInPixels()
        {
            GetDescriptor(PictureType);
            return ValidateImage(Data, PictureType);
        }

        internal XMLSlideShow GetSlideShow()
        {
            for (var parent = GetParent(); parent != null; parent = parent.GetParent())
            {
                if (parent is XMLSlideShow presentation)
                {
                    return presentation;
                }
            }
            throw new POIXMLException("The image is not attached to a presentation.");
        }

        internal void EnsureAttached()
        {
            var presentation = GetSlideShow();
            presentation.EnsureOpen();
            var part = GetPackagePart();
            if (!ReferenceEquals(part.Package, presentation.Package) ||
                !ReferenceEquals(presentation.Package.GetPart(part.PartName), part))
            {
                throw new InvalidOperationException("The image part is no longer attached to this presentation.");
            }
        }

        internal static bool IsImageContentType(string contentType)
        {
            return contentType != null && contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        }

        internal static XSLFRelation GetDescriptor(PictureType format)
        {
            switch (format)
            {
                case PictureType.PNG: return XSLFRelation.IMAGE_PNG;
                case PictureType.JPEG: return XSLFRelation.IMAGE_JPEG;
                default: throw new NotSupportedException("Only embedded PNG and JPEG picture creation is supported.");
            }
        }

        internal static SKSizeI ValidateImage(byte[] data, PictureType format)
        {
            GetDescriptor(format);
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }
            using var input = new MemoryStream(data, false);
            using var codec = SKCodec.Create(input);
            if (codec == null || codec.EncodedFormat !=
                (format == PictureType.PNG ? SKEncodedImageFormat.Png : SKEncodedImageFormat.Jpeg))
            {
                throw new ArgumentException("The image bytes do not match the declared PNG/JPEG format.", nameof(data));
            }
            if (format == PictureType.PNG && !HasPngTerminator(data))
            {
                throw new ArgumentException("A complete PNG chunk sequence is required.", nameof(data));
            }
            using var bitmap = new SKBitmap();
            if (codec.Info.Width <= 0 || codec.Info.Height <= 0 || !bitmap.TryAllocPixels(codec.Info))
            {
                throw new ArgumentException("The image pixel dimensions cannot be decoded.", nameof(data));
            }
            if (codec.GetPixels(bitmap.Info, bitmap.GetPixels()) != SKCodecResult.Success)
            {
                throw new ArgumentException("A complete, decodable PNG/JPEG image is required.", nameof(data));
            }
            return new SKSizeI(codec.Info.Width, codec.Info.Height);
        }

        private static bool HasPngTerminator(byte[] data)
        {
            // Pixel decoding can succeed before the final PNG chunk has been read.
            for (var offset = 8; offset <= data.Length - 12;)
            {
                var length = ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) |
                    ((uint)data[offset + 2] << 8) | data[offset + 3];
                if (length > data.Length - offset - 12)
                {
                    return false;
                }
                if (data[offset + 4] == 'I' && data[offset + 5] == 'E' &&
                    data[offset + 6] == 'N' && data[offset + 7] == 'D')
                {
                    return length == 0;
                }
                offset += (int)length + 12;
            }
            return false;
        }

        protected internal override void PrepareForCommit()
        {
            // Image bytes live in the part and must not be cleared by POIXML save traversal.
        }
    }
}
