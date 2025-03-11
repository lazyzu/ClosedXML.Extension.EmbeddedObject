using System;
using System.IO;

namespace ClosedXML.Extension.EmbeddedObject.Appender.Icon
{
    public static class IconSet
    {
        private static readonly Lazy<string> _namespace;
        private static readonly Lazy<System.Reflection.Assembly> _assembly;

        static IconSet()
        {
            _namespace = new Lazy<string>(() =>
            {
                return typeof(IconSet).Namespace;
            });

            _assembly = new Lazy<System.Reflection.Assembly>(() =>
            {
                return typeof(IconSet).Assembly;
            });
        }

        public static class Html
        {
            public static Stream GetStream() => getEmbeddedResourceIcon("chrome.emf");    // https://icon-sets.iconify.design/logos/chrome/, html.emf
            public static EmbeddedObjectInfo.IconInfo GetIconInfo() => new EmbeddedObjectInfo.IconInfo
            {
                ImageContentType = "image/x-emf",
                ImageStreamGetter = GetStream
            };
        }

        public static class Pdf
        {
            public static Stream GetStream() => getEmbeddedResourceIcon("pdf.emf"); // https://icon-sets.iconify.design/vscode-icons/file-type-pdf2/
            public static EmbeddedObjectInfo.IconInfo GetIconInfo() => new EmbeddedObjectInfo.IconInfo
            {
                ImageContentType = "image/x-emf",
                ImageStreamGetter = GetStream
            };
        }

        public static class Txt
        {
            public static Stream GetStream() => getEmbeddedResourceIcon("text.emf");    // https://icon-sets.iconify.design/vscode-icons/file-type-text/
            public static EmbeddedObjectInfo.IconInfo GetIconInfo() => new EmbeddedObjectInfo.IconInfo
            {
                ImageContentType = "image/x-emf",
                ImageStreamGetter = GetStream
            };
        }

        public static class Zip
        {
            public static Stream GetStream() => getEmbeddedResourceIcon("zip.emf");    // https://icon-sets.iconify.design/openmoji/winrar/
            public static EmbeddedObjectInfo.IconInfo GetIconInfo() => new EmbeddedObjectInfo.IconInfo
            {
                ImageContentType = "image/x-emf",
                ImageStreamGetter = GetStream
            };
        }

        public static class Chart
        {
            public static Stream GetStream() => getEmbeddedResourceIcon("chart.png");    // https://icon-sets.iconify.design/emojione-v1/stock-chart/
            public static EmbeddedObjectInfo.IconInfo GetIconInfo() => new EmbeddedObjectInfo.IconInfo
            {
                ImageContentType = "image/png",
                ImageStreamGetter = GetStream
            };
        }

        // All icon is convert by https://cloudconvert.com/svg-to-emf, size: 256*256

        private static Stream getEmbeddedResourceIcon(string fileName)
        {
            return _assembly.Value.GetManifestResourceStream($"{_namespace.Value}.{fileName}");
        }
    }
}
