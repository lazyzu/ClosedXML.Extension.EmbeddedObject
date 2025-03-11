using ClosedXML.Extension.EmbeddedObject.Appender.Sub;
using NUnit.Framework;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject.Validation.Appender
{
    public class VmlDrawingTest
    {
        [Test]
        public async Task InitVmlDrawingXml()
        {
            var missingNamespaceByteArray = Encoding.ASCII.GetBytes(@"<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">
    <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
        <v:stroke joinstyle=""miter""/>
        <v:formulas>
            <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
            <v:f eqn=""sum @0 1 0""/>
            <v:f eqn=""sum 0 0 @1""/>
            <v:f eqn=""prod @2 1 2""/>
            <v:f eqn=""prod @3 21600 pixelWidth""/>
            <v:f eqn=""prod @3 21600 pixelHeight""/>
            <v:f eqn=""sum @0 0 1""/>
            <v:f eqn=""prod @6 1 2""/>
            <v:f eqn=""prod @7 21600 pixelWidth""/>
            <v:f eqn=""sum @8 21600 0""/>
            <v:f eqn=""prod @7 21600 pixelHeight""/>
            <v:f eqn=""sum @10 21600 0""/>
        </v:formulas>
        <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
        <o:lock v:ext=""edit"" aspectratio=""t""/>
    </v:shapetype>
</xml>");
            var missingShapetypeByteArray = Encoding.ASCII.GetBytes(@"<xml xmlns:v=""urn:schemas-microsoft-com:vml""
    xmlns:o=""urn:schemas-microsoft-com:office:office""
    xmlns:x=""urn:schemas-microsoft-com:office:excel"">
    <v:shapetype id=""_x0000_t76"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
        <v:stroke joinstyle=""miter""/>
        <v:formulas>
            <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
            <v:f eqn=""sum @0 1 0""/>
            <v:f eqn=""sum 0 0 @1""/>
            <v:f eqn=""prod @2 1 2""/>
            <v:f eqn=""prod @3 21600 pixelWidth""/>
            <v:f eqn=""prod @3 21600 pixelHeight""/>
            <v:f eqn=""sum @0 0 1""/>
            <v:f eqn=""prod @6 1 2""/>
            <v:f eqn=""prod @7 21600 pixelWidth""/>
            <v:f eqn=""sum @8 21600 0""/>
            <v:f eqn=""prod @7 21600 pixelHeight""/>
            <v:f eqn=""sum @10 21600 0""/>
        </v:formulas>
        <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
        <o:lock v:ext=""edit"" aspectratio=""t""/>
    </v:shapetype>
</xml>");

            using (var emptyStream = new MemoryStream())
            using (var missingNamespaceStream = new MemoryStream())
            using (var missingShapetypeStream = new MemoryStream())
            {
                await missingNamespaceStream.WriteAsync(missingNamespaceByteArray);
                await missingShapetypeStream.WriteAsync(missingShapetypeByteArray);

                VmlDrawingAppender.InitVmlDrawingXml(emptyStream);
                VmlDrawingAppender.InitVmlDrawingXml(missingNamespaceStream);
                VmlDrawingAppender.InitVmlDrawingXml(missingShapetypeStream);

                var emptyStreamHandleResult = Encoding.ASCII.GetString(emptyStream.ToArray());
                var missingNamespaceStreamHandleResult = Encoding.ASCII.GetString(missingNamespaceStream.ToArray());
                var missingShapetypeStreamHandleResult = Encoding.ASCII.GetString(missingShapetypeStream.ToArray());

                Assert.That(emptyStreamHandleResult, Is.EqualTo(@"???<xml xmlns:v=""urn:schemas-microsoft-com:vml""
    xmlns:o=""urn:schemas-microsoft-com:office:office""
    xmlns:x=""urn:schemas-microsoft-com:office:excel"">
    <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
        <v:stroke joinstyle=""miter""/>
        <v:formulas>
            <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
            <v:f eqn=""sum @0 1 0""/>
            <v:f eqn=""sum 0 0 @1""/>
            <v:f eqn=""prod @2 1 2""/>
            <v:f eqn=""prod @3 21600 pixelWidth""/>
            <v:f eqn=""prod @3 21600 pixelHeight""/>
            <v:f eqn=""sum @0 0 1""/>
            <v:f eqn=""prod @6 1 2""/>
            <v:f eqn=""prod @7 21600 pixelWidth""/>
            <v:f eqn=""sum @8 21600 0""/>
            <v:f eqn=""prod @7 21600 pixelHeight""/>
            <v:f eqn=""sum @10 21600 0""/>
        </v:formulas>
        <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
        <o:lock v:ext=""edit"" aspectratio=""t""/>
    </v:shapetype>
</xml>"));

                Assert.That(missingNamespaceStreamHandleResult, Is.EqualTo(@"???<?xml version=""1.0"" encoding=""utf-8""?>
<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
  <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
    <v:stroke joinstyle=""miter"" />
    <v:formulas>
      <v:f eqn=""if lineDrawn pixelLineWidth 0"" />
      <v:f eqn=""sum @0 1 0"" />
      <v:f eqn=""sum 0 0 @1"" />
      <v:f eqn=""prod @2 1 2"" />
      <v:f eqn=""prod @3 21600 pixelWidth"" />
      <v:f eqn=""prod @3 21600 pixelHeight"" />
      <v:f eqn=""sum @0 0 1"" />
      <v:f eqn=""prod @6 1 2"" />
      <v:f eqn=""prod @7 21600 pixelWidth"" />
      <v:f eqn=""sum @8 21600 0"" />
      <v:f eqn=""prod @7 21600 pixelHeight"" />
      <v:f eqn=""sum @10 21600 0"" />
    </v:formulas>
    <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect"" />
    <o:lock v:ext=""edit"" aspectratio=""t"" />
  </v:shapetype>
</xml>"));

                Assert.That(missingNamespaceStreamHandleResult, Is.EqualTo(@"???<?xml version=""1.0"" encoding=""utf-8""?>
<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
  <v:shapetype id=""_x0000_t76"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
    <v:stroke joinstyle=""miter"" />
    <v:formulas>
      <v:f eqn=""if lineDrawn pixelLineWidth 0"" />
      <v:f eqn=""sum @0 1 0"" />
      <v:f eqn=""sum 0 0 @1"" />
      <v:f eqn=""prod @2 1 2"" />
      <v:f eqn=""prod @3 21600 pixelWidth"" />
      <v:f eqn=""prod @3 21600 pixelHeight"" />
      <v:f eqn=""sum @0 0 1"" />
      <v:f eqn=""prod @6 1 2"" />
      <v:f eqn=""prod @7 21600 pixelWidth"" />
      <v:f eqn=""sum @8 21600 0"" />
      <v:f eqn=""prod @7 21600 pixelHeight"" />
      <v:f eqn=""sum @10 21600 0"" />
    </v:formulas>
    <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect"" />
    <o:lock v:ext=""edit"" aspectratio=""t"" />
  </v:shapetype>
  <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" filled=""false"" stroked=""false"" o:spt=""75"" o:preferrelative=""true"" path=""m@4@5l@4@11@9@11@9@5xe"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:v=""urn:schemas-microsoft-com:vml"">
    <v:stroke joinstyle=""miter"" />
    <v:formulas>
      <v:f eqn=""if lineDrawn pixelLineWidth 0"" />
      <v:f eqn=""sum @0 1 0"" />
      <v:f eqn=""sum 0 0 @1"" />
      <v:f eqn=""prod @2 1 2"" />
      <v:f eqn=""prod @3 21600 pixelWidth"" />
      <v:f eqn=""prod @3 21600 pixelHeight"" />
      <v:f eqn=""sum @0 0 1"" />
      <v:f eqn=""prod @6 1 2"" />
      <v:f eqn=""prod @7 21600 pixelWidth"" />
      <v:f eqn=""sum @8 21600 0"" />
      <v:f eqn=""prod @7 21600 pixelHeight"" />
      <v:f eqn=""sum @10 21600 0"" />
    </v:formulas>
    <v:path gradientshapeok=""true"" o:connecttype=""rect"" o:extrusionok=""false"" />
    <o:lock v:ext=""edit"" aspectratio=""true"" />
  </v:shapetype>
</xml>"));
            }
        }
    }
}
