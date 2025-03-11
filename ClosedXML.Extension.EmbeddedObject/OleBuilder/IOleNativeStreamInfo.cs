using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject.OleBuilder
{
    // reference: https://github.com/idiom/OLEPackagerFormat

    public interface IOleNativeStreamInfo
    {
        string FileName { get; }
        string FilePath { get; }
        string TemporaryPath { get; }
        Stream Data { get; }
    }

    public class OleNativeStreamInfo : IOleNativeStreamInfo
    {
        public string FileName { get; init; }
        public string FilePath { get; init; }
        public string TemporaryPath { get; init; }
        public Stream Data { get; init; }
    }

    public static class OleNativeStreamBuilder
    {
        public static Stream ToStream(this IOleNativeStreamInfo objStreamInfo)
        {
            LoadOleNativeContentParts(objStreamInfo
                , out var header
                , out var fileName
                , out var filePath
                , out var utype
                , out var temporaryPath
                , out var dataLength
                , out var temporaryPathUnicode
                , out var fileNameUnicode
                , out var filePathUnicode
                , out var packageLength);

            var result = new MemoryStream();
            using (BinaryWriter writer = new BinaryWriter(result, Encoding.UTF8, leaveOpen: true))
            {
                writer.Write(BitConverter.GetBytes((uint)packageLength));
                writer.Write(header);
                writer.Write(fileName);
                writer.Write(filePath);
                writer.Write(utype);
                writer.Write(temporaryPath);

                writer.Write(dataLength);
                objStreamInfo.Data.Position = 0;
                objStreamInfo.Data.CopyTo(result);

                writer.Write(temporaryPathUnicode);
                writer.Write(fileNameUnicode);
                writer.Write(filePathUnicode);
            }

            result.Position = 0;
            return result;
        }

        public static async Task<Stream> ToStreamAsync(this IOleNativeStreamInfo objStreamInfo, CancellationToken cancellationToken = default)
        {
            LoadOleNativeContentParts(objStreamInfo
                , out var header
                , out var fileName
                , out var filePath
                , out var utype
                , out var temporaryPath
                , out var dataLength
                , out var temporaryPathUnicode
                , out var fileNameUnicode
                , out var filePathUnicode
                , out var packageLength);

            var result = new MemoryStream();
            using (BinaryWriter writer = new BinaryWriter(result, Encoding.UTF8, leaveOpen: true))
            {
                writer.Write(BitConverter.GetBytes((uint)packageLength));
                writer.Write(header);
                writer.Write(fileName);
                writer.Write(filePath);
                writer.Write(utype);
                writer.Write(temporaryPath);

                writer.Write(dataLength);
                objStreamInfo.Data.Position = 0;
                await objStreamInfo.Data.CopyToAsync(result, 81920  /* default value */, cancellationToken);

                writer.Write(temporaryPathUnicode);
                writer.Write(fileNameUnicode);
                writer.Write(filePathUnicode);
            }

            result.Position = 0;
            return result;
        }

        public static void LoadOleNativeContentParts(IOleNativeStreamInfo objStreamInfo
            , out byte[] header
            , out byte[] fileName
            , out byte[] filePath
            , out byte[] utype
            , out byte[] temporaryPath
            , out byte[] dataLength
            , out byte[] temporaryPathUnicode
            , out byte[] fileNameUnicode
            , out byte[] filePathUnicode
            , out long packageLength)
        {
            var absoluteFilePath = objStreamInfo.FilePath;
            var absoluteTempPath = objStreamInfo.TemporaryPath;

            header = new byte[] { 0x02, 0x00 };
            fileName = ToNullTerminateStream(objStreamInfo.FileName);
            filePath = ToNullTerminateStream(absoluteFilePath);
            utype = new byte[] { 0x00, 0x00, 0x03, 0x00 };
            temporaryPath = ToLengthPrefixedAnsiString(absoluteTempPath);
            dataLength = ToDataLengthStream(objStreamInfo.Data);
            temporaryPathUnicode = ToLengthPrefixedUnicodeString(absoluteTempPath);
            fileNameUnicode = ToLengthPrefixedUnicodeString(objStreamInfo.FileName);
            filePathUnicode = ToLengthPrefixedUnicodeString(absoluteFilePath);
            packageLength = header.Length + fileName.Length
                                              + filePath.Length
                                              + utype.Length
                                              + temporaryPath.Length
                                              + 4 + objStreamInfo.Data.Length
                                              + fileNameUnicode.Length
                                              + temporaryPathUnicode.Length
                                              + filePathUnicode.Length;
        }

        private static byte[] ToDataLengthStream(Stream data)
        {
            if (data == null || data.Length == 0) throw new ArgumentException(nameof(data));
            else
            {
                var dataLengthNumber = (uint)data.Length;
                return BitConverter.GetBytes(dataLengthNumber);
            }
        }

        public static byte[] ToNullTerminateStream(string source)
        {
            var _source = source?.Trim();
            if (string.IsNullOrEmpty(_source)) throw new ArgumentException(nameof(source));
            else
            {
                return Encoding.ASCII.GetBytes(_source).Concat(new byte[]
                {
                    0x00
                }).ToArray();
            }
        }

        public static byte[] ToLengthPrefixedAnsiString(string source)
        {
            var _source = source?.Trim();

            if (string.IsNullOrEmpty(_source)) throw new ArgumentException(nameof(_source));
            else
            {
                var lengthPrefixNumber = _source.Length + 1;
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

        public static byte[] ToLengthPrefixedUnicodeString(string source)
        {
            var _source = source?.Trim();

            if (string.IsNullOrEmpty(_source)) throw new ArgumentException(nameof(_source));
            else
            {
                var lengthPrefixNumber = (uint)_source.Length;
                var resultStreamLength = 4 + lengthPrefixNumber * 2;
                var resultStream = new byte[resultStreamLength];

                // Length Prefix
                var lengthPrefixByteArray = BitConverter.GetBytes(lengthPrefixNumber);
                for (int i = 0; i < 4; i++) resultStream[i] = lengthPrefixByteArray[i];

                // Ansi String
                var ansiStringByteArray = Encoding.Unicode.GetBytes(_source);
                for (int i = 0; i < ansiStringByteArray.Length; i++) resultStream[i + 4] = ansiStringByteArray[i];

                return resultStream;
            }
        }

    }
}
