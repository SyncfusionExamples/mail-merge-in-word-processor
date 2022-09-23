using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using System;
using System.IO;
using Syncfusion.EJ2.DocumentEditor;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;

namespace MailMergeExample
{
    [Route("api/DocumentEditor")]
    [ApiController]
    public class DocumentEditorController : ControllerBase
    {
#if NET6_0
        private readonly IWebHostEnvironment _hostingEnvironment;
        public DocumentEditorController(IWebHostEnvironment hostingEnvironment)
#else
        private readonly IHostingEnvironment _hostingEnvironment;
        public DocumentEditorController(IHostingEnvironment hostingEnvironment)
#endif
        {
            _hostingEnvironment = hostingEnvironment;
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("Default")]
        public string Default()
        {
            return "Default";
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("Import")]
        public string Import(IFormCollection data)
        {

            Stream stream = new MemoryStream();
            string type = "docx";
            if (data.Files.Count == 0)
                return null;
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            type = index > -1 && index < file.FileName.Length - 1 ? file.FileName.Substring(index + 1) : "";
            file.CopyTo(stream);
            stream.Position = 0;

            WordDocument document = WordDocument.Load(stream, GetFormatType(type.ToLower()));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody] CustomParameter param)
        {
            if (param.content != null && param.content != "")
            {
                WordDocument document = WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
                return json;
            }
            return "";
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("MailMergeReport")]
        public string MailMergeReport([FromBody] ExportData exportData)
        {
            Byte[] data = Convert.FromBase64String(exportData.documentData.Split(',')[1]);
            MemoryStream stream = new MemoryStream();
            stream.Write(data, 0, data.Length);
            stream.Position = 0;
            List<string> sfdtFiles = new List<string>();
            if (string.IsNullOrEmpty(exportData.mergeData))
            {
                return "";
            }
            Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
            stream.Dispose();
            if (string.IsNullOrEmpty(exportData.recordCount))
            {
                document.ExecuteMailMerge(exportData.mergeData);
                Syncfusion.EJ2.DocumentEditor.WordDocument wordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(document);
                string sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(wordDocument);
                wordDocument.Dispose();
                return sfdtText;
            }
            else
            {
                int recordCount = int.Parse(exportData.recordCount);
                for (int i = 0; i < recordCount; i++)
                {
                    Syncfusion.DocIO.DLS.WordDocument templateDocument = document.Clone();
                    string mergeData = exportData.mergeData.Replace("\"RecordId\":\"-1\"", "\"RecordId\":\"" + i.ToString() + "\"");
                    templateDocument.ExecuteMailMerge(mergeData);
                    Syncfusion.EJ2.DocumentEditor.WordDocument wordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(templateDocument);
                    string sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(wordDocument);
                    sfdtFiles.Add(sfdtText);
                    wordDocument.Dispose();
                }
                string x = Newtonsoft.Json.JsonConvert.SerializeObject(sfdtFiles);
                return x;
            }
        }
        public class ExportData
        {
            public string fileName { get; set; }
            public string documentData { get; set; }
            public string mergeData { get; set; }
            public string recordCount { get; set; }
        }


        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("RestrictEditing")]
        public string[] RestrictEditing([FromBody] CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }

        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            switch (format.ToLower())
            {
                case "dotx":
                case "docx":
                case "docm":
                case "dotm":
                    return FormatType.Docx;
                case "dot":
                case "doc":
                    return FormatType.Doc;
                case "rtf":
                case ".rtf":
                    return FormatType.Rtf;
                case "txt":
                    return FormatType.Txt;
                case "xml":
                    return FormatType.WordML;
                case "html":
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            }
        }

        internal WDocument GetDocument(IFormCollection data)
        {
            Stream stream = new MemoryStream();
            if (data.Files.Count == 0)
                return null;
            IFormFile file = data.Files[0];

            file.CopyTo(stream);
            stream.Position = 0;

            WDocument document = new WDocument(stream, WFormatType.Docx);
            stream.Dispose();
            return document;
        }
    }

    public class CustomParameter
    {
        public string content { get; set; }
        public string type { get; set; }
    }

    public class CustomRestrictParameter
    {
        public string passwordBase64 { get; set; }
        public string saltBase64 { get; set; }
        public int spinCount { get; set; }
    }
}
