using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace ConsoleApplication
{
    internal class Program
    {
        private static object powerPoint;
        private static object presentations;
        private const int FileTypePPTX = 24;

        static void Main(string[] args)
        {
            var dir = @"C:\Users\andyedinborough\Downloads\";
            var jsonFile = @"C:\temp\hymnal.json";
            ConvertAllPptToPptx(dir);
            GenerateHymnal(dir, jsonFile);
        }

        static void GenerateHymnal(string dir, string jsonFile)
        {
            var temp = System.IO.Path.GetTempPath();
            var zipdir = System.IO.Path.Combine(temp, "pptx-zip");
            System.IO.Directory.CreateDirectory(zipdir);

            var pptxs = System.IO.Directory.GetFiles(dir, "*.pptx");

            var hymnal = pptxs.Select(pptx =>
            {
                Console.WriteLine(pptx);
                System.IO.Directory.Delete(zipdir, true);
                System.IO.Directory.CreateDirectory(zipdir);

                ZipFile.ExtractToDirectory(pptx, zipdir);
                var slidesDir = System.IO.Path.Combine(zipdir, "ppt\\slides");
                var slides = System.IO.Directory.GetFiles(slidesDir, "*.xml");

                return new
                {
                    file = System.IO.Path.GetFileName(pptx),
                    text = slides.Select(slide =>
                    {
                        var xdoc = XDocument.Load(slide);
                        var a = XNamespace.Get("http://schemas.openxmlformats.org/drawingml/2006/main");
                        var t = xdoc.Descendants(a + "t");

                        return t.Select(elm => elm.Value)
                            .Where(x => x.Length > 0)
                            .ToList();
                    })
                    .Where(x => x.Count > 0)
                    .ToList()
                };

            }).ToList();

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(hymnal, Newtonsoft.Json.Formatting.Indented);
            System.IO.File.WriteAllText(jsonFile, json);
        }

        private static void ConvertAllPptToPptx(string dir)
        {
            try
            {
                powerPoint = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
                powerPoint.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, powerPoint, new object[] { true });
                presentations = powerPoint.GetType().InvokeMember("Presentations", BindingFlags.GetProperty, null, powerPoint, Type.EmptyTypes);

                var ppts = Directory.GetFiles(dir, "*.ppt");
                foreach (var ppt in ppts)
                {
                    var pptx = Path.Combine(Path.GetDirectoryName(ppt), Path.GetFileNameWithoutExtension(ppt) + ".pptx");
                    ConvertPptToPptx(ppt, pptx);
                }
            }
            finally
            {
                Method(powerPoint, "Quit");
            }
        }

        private static void ConvertPptToPptx(string pptFilename, string pptxFilename)
        {
            var presentation = Method(presentations, "Open", pptFilename, true, false, false);
            Method(presentation, "SaveAs", pptxFilename, FileTypePPTX);
            Method(presentation, "Close");
        }

        private static object Method(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, target, args ?? Type.EmptyTypes);
        }
    }
}