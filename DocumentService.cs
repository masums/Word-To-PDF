using ImageMagick;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Drawing;
using Syncfusion.Licensing;
using Syncfusion.Pdf;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace Pusintek.AspNetcore.DocIO
{

    public class DocumentService
    {
        /// <summary>  
        /// <para>Library for generate pdf file from word template with mail merge</para>
        /// The <paramref name="_templateName"/> is path for document template
        /// <para>Sample Code</para>
        /// <code>
        /// DocumentService documentService = new DocumentService("wordtemplate/test.docx");
        ///<para>documentService.GeneratePDF();</para>
        /// </code>
        /// </summary> 

        public DocumentService(string _templateName)
        {
            TemplateName = _templateName;
            GetDocumentTemplate();
            SyncfusionLicenseProvider.RegisterLicense("MzUzMDVAMzEzNjJlMzMyZTMwanhURzlmMFF1MElUOXRTNHhHYWFjZExlTS8vNStNQ1FOdngvWWN3dFRjdz0=");
        }

        #region Class properties
        public string TemplateName = "";
        public DataField  Data = new DataField();
        public List<DataFieldGroup> DataTable = new List<DataFieldGroup>();
        public Dictionary<string, Options> Options;
        private WordDocument DocumentTemplate;
        private MemoryStream PDFStreamOutput;
        private PdfDocument PDFGenerated;
        private string[] fieldNames;
        private string[] fieldValues;
        private MemoryStream HTMLStreamOutput = new MemoryStream();
        private List<string> Keys = new List<string>();
        #endregion

        #region Generate Function
        /// <summary>
        /// <para>Generate and save generated file on server</para>
        /// <code>
        /// <para>DocumentService documentService = new DocumentService("test.docx");</para>
        ///  <para>documentService.GeneratePDF("pdfGenerate", "test.pdf");</para>
        /// </code>
        /// </summary>
        /// <param name="outputPath"> Path for save generated file</param>
        public void GeneratePDF(string OutPutPath, string OutputName)
        {
            AllowToProcess();
            if (OutputName == "")
                OutputName = DateTime.Now.ToString();
            PDFStreamOutput = GenerateMemoriStreamPDF();

            if (!Directory.Exists(OutPutPath))
                Directory.CreateDirectory(OutPutPath);
            using (var outputStream = File.Create($@"{OutPutPath}/{OutputName}"))
            {
                PDFStreamOutput.CopyTo(outputStream);
            }
        }

        /// <summary>
        /// Generate PDF File and get memory which contain it
        /// <returns>Memory stream which contain generated PDF(MemoryStream) </returns>
        /// </summary>
        public MemoryStream GeneratePDF()
        {
            AllowToProcess();
            return GenerateMemoriStreamPDF();
        }

        /// <summary>
        /// Generate html string from document template
        /// </summary>
        /// <returns>string of html</returns>
        public string GenerateHtml()
        {
            ProcessDocument();
            ProcessDocumentGroup();
            HTMLExport export = new HTMLExport();
            export.SaveAsXhtml(DocumentTemplate, HTMLStreamOutput);
            StreamReader reader = new StreamReader(HTMLStreamOutput);
            HTMLStreamOutput.Position = 0;
            return reader.ReadToEnd();
        }

        #endregion

        /// <summary>
        /// Release all used resource
        /// </summary>
        public void ReleaseResource()
        {
            DocumentTemplate = null;
            PDFStreamOutput = null;
            PDFGenerated = null;
            DocumentTemplate.Close();
        }

        #region Document Processing
        /// <summary>
        /// Replace Mail merge with value,
        /// </summary>
        protected void ProcessDocument()
        {
            if (Data.Data.Count() == 0)
                return;
            Keys = Data.Options.Keys.ToList();
            Options = Data.Options;
            DocumentTemplate.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ImageHandler);

            fieldValues = Data.Data.Values.Select(i => i.ToString()).ToArray();
            fieldNames= Data.Data.Keys.Select(i => i.ToString()).ToArray();
            DocumentTemplate.MailMerge.Execute(fieldNames, fieldValues);
        }

        /// <summary>
        /// Replace mail merge with list of object
        /// </summary>
        /// <param name="processedData">Dictionary with list object</param>
        protected void ProcessDocumentGroup()
        {
            if (DataTable.Count() == 0)
                return;
            List<DictionaryEntry> commands = new List<DictionaryEntry>();

            foreach (var data in DataTable)
            {
                DocumentTemplate.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ImageHandler);
                Options = data.Options;
                MailMergeDataTable dataTable = new MailMergeDataTable(data.Key, data.Data);
                Options = data.Options;                
                DocumentTemplate.MailMerge.ExecuteGroup(dataTable);
            }
        }
        #endregion

        #region generate PDF
        /// <summary>
        /// <para>Render PDF</para>
        /// </summary>
        protected void CreatePDF()
        {
            // Creates a new instance of DocIORenderer class.
            DocIORenderer render = new DocIORenderer();

            // Converts Word document into PDF document.
            PDFGenerated = render.ConvertToPDF(DocumentTemplate);
            render.Dispose();
            DocumentTemplate.Dispose();
        }

        private MemoryStream GenerateMemoriStreamPDF()
        {
            MemoryStream outputStream = new MemoryStream();
            try
            {
                ProcessDocumentGroup();
                ProcessDocument();
                CreatePDF();
                // Save the PDF document
                PDFGenerated.Save(outputStream);
                outputStream.Position = 0;
            }
            catch
            {
                throw;
            }
            return outputStream;
        }
        #endregion

        private void GetDocumentTemplate()
        {
            FileStream templateStream = new FileStream($@"{TemplateName}", FileMode.Open, FileAccess.Read);
            
            DocumentTemplate = new WordDocument(templateStream, FormatType.Docx);
            templateStream.Dispose();
            templateStream = null;
        }

        private void AllowToProcess()
        {
            if (TemplateName == "") throw new Exception("Template belum didefinisikan");
        }

        private string ExtFile()
        {
            var ArrayName = TemplateName.Split('.');
            return ArrayName[ArrayName.Length - 1];
        }

        #region Modify Value if using images
        private void MergeField_ImageHandler(object sender, MergeImageFieldEventArgs args)
        {
            args.ImageStream = new MemoryStream();
            if (Options.Keys.Contains( args.FieldName))
            {
                if (args.FieldValue == null)
                    return;
                Options options = Options[args.FieldName.ToString()];               
                try
                {                        
                    args.ImageStream = GetImageStream(args.FieldValue.ToString(), options);         
                }
                catch(Exception ex)
                {
                    args.Text = $"Failed to fetch image {args.FieldValue}";
                }
            }
        }

        private Stream GetImageStream (string path, Options options)
        {
            MemoryStream stream = new MemoryStream();
            if (options.FromUrl)
            {
                WebClient httpClient = new WebClient();
                byte[] result = httpClient.DownloadData(path);
                
                stream.Write(result, 0, result.Length);
            }
            else
            {
                using (var fileStream = File.Open($@"{path}", FileMode.Open))
                {
                    fileStream.CopyTo(stream);
                }
            }
            return ResizeImage(stream, options);
        }

        private Stream ResizeImage(MemoryStream stream, Options options)
        {
            Stream ResizedStream = new MemoryStream();
            stream.Position = 0;
            MagickImage image = new MagickImage(stream);
            
            if (options.PercentageResize != 0)
            {
                Percentage percentage = new Percentage(options.PercentageResize);
                image.Resize(percentage);
            }
            else
            {
                int height = image.Height;
                int width = image.Width;
                if (options.Height != 0)
                    height = options.Height;
                if (options.Width != 0)
                    width = options.Width;
                image.Resize(width, height);
            };
            byte[] imageByte = image.ToByteArray();
            ResizedStream.Position = 0;
            ResizedStream.Write(imageByte, 0, imageByte.Length);
            
            return ResizedStream; 
        }

        #endregion
    }
}
