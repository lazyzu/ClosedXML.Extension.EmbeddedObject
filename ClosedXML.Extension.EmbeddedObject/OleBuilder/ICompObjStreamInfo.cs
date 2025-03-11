using System;
using System.IO;
using System.Linq;
using System.Text;

namespace ClosedXML.Extension.EmbeddedObject.OleBuilder
{
    // Reference: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleds/142e0420-2f74-4ed9-829b-0b3d5a684d01

    public interface ICompObjStreamInfo
    {
        byte[] Header { get; }

        string AnsiUserType { get; }

        OleClipboardFormat AnsiClipboardFormat { get; }

        byte[] Reserved1 { get; }

        // Not support Unicode: UnicodeUserType, UnicodeClipboardFormat, Reserved2
    }

    public class SimpleCompObjStreamInfo : ICompObjStreamInfo
    {
        public byte[] Header => new byte[]
        {
            0x01,   // Reserved 1
            0x00,
            0xFE,
            0xFF,
            0x03,   // Version
            0x0A,
            0x00,
            0x00,
            0xFF,   // Reserved 2
            0xFF,
            0xFF,
            0xFF,
            0x0C,
            0x00,
            0x03,
            0x00,
            0x00,
            0x00,
            0x00,
            0x00,
            0xC0,
            0x00,
            0x00,
            0x00,
            0x00,
            0x00,
            0x00,
            0x46
        };

        public string AnsiUserType => "OLE Package";

        public OleClipboardFormat AnsiClipboardFormat => OleClipboardFormat.Registered;

        public byte[] Reserved1 => CompObjStreamBuilder.ToLengthPrefixedAnsiString("Package");
    }

    /// <summary>
    ///     The OLE version 1.0 and 2.0 clipboard formats
    /// </summary>
    public enum OleClipboardFormat
    {
        /// <summary>
        ///     The format is a registered clipboard format
        /// </summary>
        Registered = 0x00000000,

        // ReSharper disable InconsistentNaming
        /// <summary>
        ///     Bitmap16 Object structure
        /// </summary>
        CF_BITMAP = 0x00000002,

        /// <summary>
        /// </summary>
        CF_METAFILEPICT = 0x00000003,

        /// <summary>
        ///     DeviceIndependentBitmap Object structure
        /// </summary>
        CF_DIB = 0x00000008,

        /// <summary>
        ///     Enhanced Metafile
        /// </summary>
        CF_ENHMETAFILE = 0x0000000E
        // ReSharper restore InconsistentNaming
    }

    public static class CompObjStreamBuilder
    {
        public static Stream ToStream(this ICompObjStreamInfo objStreamInfo)
        {
            if (objStreamInfo == null) throw new ArgumentNullException(nameof(objStreamInfo));
            if (objStreamInfo.Header == null) throw new ArgumentNullException(nameof(objStreamInfo.Header));
            if (objStreamInfo.Reserved1 == null) throw new ArgumentNullException(nameof(objStreamInfo.Reserved1));

            var header = objStreamInfo.Header;
            var ansiUserType = ansiUserTypeToByteStream(objStreamInfo.AnsiUserType);
            var ansiClipboardFormat = ansiClipboardFormatToByteStream(objStreamInfo.AnsiClipboardFormat);
            var reserved1 = objStreamInfo.Reserved1;
            var unicodeMarker = UnicodeMarker;
            var unicodeUserType = Empty;
            var unicodeClipboardFormat = Empty;
            var reserved2 = Empty;

            var result = new MemoryStream();
            using (BinaryWriter writer = new BinaryWriter(result, Encoding.UTF8, leaveOpen: true))
            {
                writer.Write(ansiUserType);
                writer.Write(ansiClipboardFormat);
                writer.Write(reserved1);
                writer.Write(unicodeMarker);
                writer.Write(unicodeUserType);
                writer.Write(unicodeClipboardFormat);
                writer.Write(reserved2);
            }
            result.Position = 0;
            return result;
        }

        private static byte[] ansiUserTypeToByteStream(string ansiUserType)
        {
            return ToLengthPrefixedAnsiString(ansiUserType);
        }

        private static byte[] ansiClipboardFormatToByteStream(OleClipboardFormat clipboardFormat)
        {
            if (clipboardFormat == OleClipboardFormat.Registered) return Empty;
            else
            {
                // 0xffffffff or 0xfffffffe
                return new byte[] { 0xff, 0xff, 0xff, 0xff }.Concat(BitConverter.GetBytes((uint)clipboardFormat)).ToArray();
            }
        }

        public static byte[] ToLengthPrefixedAnsiString(string source)
        {
            var _source = source?.Trim();

            if (string.IsNullOrEmpty(_source)) throw new ArgumentException(nameof(_source));
            else
            {
                var lengthPrefixNumber = (uint)(_source.Length + 1);
                var resultStreamLength = 4 + lengthPrefixNumber;
                var resultStream = new byte[resultStreamLength];

                // Length Prefix
                var lengthPrefixByteArray = BitConverter.GetBytes(lengthPrefixNumber);
                for (int i = 0; i < 4; i++) resultStream[i] = lengthPrefixByteArray[i];

                // Ansi String
                var ansiStringByteArray = Encoding.ASCII.GetBytes(_source);
                for (int i = 0; i < ansiStringByteArray.Length; i++) resultStream[i + 4] = ansiStringByteArray[i];

                return resultStream;
            }
        }

        public static readonly byte[] Empty = new byte[4];

        private static readonly byte[] UnicodeMarker = new byte[]   // 0x71B239F4
        {
            0xF4,
            0x39,
            0xB2,
            0x71,
        };
    }
}
