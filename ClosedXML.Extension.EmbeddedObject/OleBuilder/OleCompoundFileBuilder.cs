using OpenMcdf;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject.OleBuilder
{
    public static class OleCompoundFileBuilder
    {
        // Support async, but async only for file content read, OpenMcdf is not support async operation so not able to implement write async operation

        public static Stream Build(ICompObjStreamInfo compObjInfo, IOleNativeStreamInfo oleNativeInfo)
        {
            var cf = new CompoundFile(CFSVersion.Ver_3, CFSConfiguration.Default);
            cf.RootStorage.CLSID = new Guid("0003000c-0000-0000-c000-000000000046");

            var compDataStream = cf.RootStorage.AddStream("\u0001CompObj");
            compDataStream.CopyFrom(compObjInfo.ToStream());

            var oleNativeStream = cf.RootStorage.AddStream("\u0001Ole10Native");
            oleNativeStream.CopyFrom(oleNativeInfo.ToStream());

            var result = new MemoryStream();
            cf.Save(result);
            result.Position = 0;
            return result;
        }

        public static async Task<Stream> BuildAsync(ICompObjStreamInfo compObjInfo, IOleNativeStreamInfo oleNativeInfo, CancellationToken cancellationToken = default)
        {
            var cf = new CompoundFile(CFSVersion.Ver_3, CFSConfiguration.Default);
            cf.RootStorage.CLSID = new Guid("0003000c-0000-0000-c000-000000000046");

            var compDataStream = cf.RootStorage.AddStream("\u0001CompObj");
            compDataStream.CopyFrom(compObjInfo.ToStream());

            var oleNativeStream = cf.RootStorage.AddStream("\u0001Ole10Native");
            oleNativeStream.CopyFrom(await oleNativeInfo.ToStreamAsync(cancellationToken));

            var result = new MemoryStream();
            cf.Save(result);
            result.Position = 0;
            return result;
        }

        public static void Build(ICompObjStreamInfo compObjInfo, IOleNativeStreamInfo oleNativeInfo, string path)
        {
            using (var oleStream = Build(compObjInfo, oleNativeInfo))
            using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
            {
                oleStream.CopyTo(outputFileStream);
            }
        }

        public static async Task BuildAsync(ICompObjStreamInfo compObjInfo, IOleNativeStreamInfo oleNativeInfo, string path, CancellationToken cancellationToken = default)
        {
            using (var oleStream = await BuildAsync(compObjInfo, oleNativeInfo, cancellationToken))
            using (FileStream outputFileStream = new FileStream(path, FileMode.Create))
            {
                await oleStream.CopyToAsync(outputFileStream, 81920 /* default value */, cancellationToken);
            }
        }
    }
}
