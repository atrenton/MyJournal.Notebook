using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using MyJournal.Notebook.Diagnostics;
using Svg;

namespace MyJournal.Notebook.Properties
{
    // Svg Package Dependency: Svg.NET
    // REF: https://www.nuget.org/packages/Svg/3.4.4
    // API: https://svg-net.github.io/SVG/api/Svg.html

    static internal class Asset
    {
        const string AssetsFolderName = "Assets";

        internal static readonly string
            Namespace = typeof(Asset).Namespace;

#nullable enable

        /// <summary>
        /// Returns Base64 encoded image data for the SVG document;
        /// ImageFormat defaults to Png.
        /// </summary>
        /// <param name="document">SVG document reference</param>
        /// <param name="format">Specifies the format of the image;
        /// defaults to PNG</param>
        internal static string GetBase64Data(
            SvgDocument document, ImageFormat? format = null)
        {
            var imageFormat = format ?? ImageFormat.Png;

            var bitmap = document.Draw();
            using var ms = new MemoryStream();
            bitmap.Save(ms, imageFormat);
            return Convert.ToBase64String(ms.ToArray());
        }

#nullable disable

        /// <summary>
        /// Returns Base64 encoded image data for the embedded resource.
        /// </summary>
        /// <param name="imageName">Image file name</param>
        internal static string GetBase64Image(string imageName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"{Namespace}.{AssetsFolderName}.{imageName}";

            using var stream = assembly.GetManifestResourceStream(resourceName);
            using var image = Image.FromStream(stream);
            using var ms = new MemoryStream();
            image.Save(ms, image.RawFormat);
            return Convert.ToBase64String(ms.ToArray());
        }

        /// <summary>
        /// Returns embedded SVG document resource.
        /// </summary>
        /// <param name="document">SVG document file name</param>
        internal static SvgDocument LoadSvgDocument(string document)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = $"{Namespace}.{AssetsFolderName}.{document}";
            using var stream = assembly.GetManifestResourceStream(resourceName);
            SvgDocument.DisableDtdProcessing = true;
            return SvgDocument.Open<SvgDocument>(stream);
        }

        /// <summary>
        /// Writes the SVG element XML content to the trace listeners.
        /// </summary>
        /// <param name="element">SVG element reference</param>
        [Conditional("DEBUG")]
        internal static async void DumpSvgElementAsync(SvgElement element)
        {
            await Task.Run(() =>
            {
                var builder = new StringBuilder();
                var settings = new XmlWriterSettings
                {
                    OmitXmlDeclaration = true,
                    Indent = true,
                    NewLineOnAttributes = true
                };

                using (var writer = XmlWriter.Create(builder, settings))
                {
                    element.Write(writer);
                }

                var lines = builder.ToString().Split(Environment.NewLine);
                foreach (var line in lines)
                {
                    Tracer.WriteDebugLine("SVG> {0}", line);
                }

            }).ConfigureAwait(false);
        }
    }
}
