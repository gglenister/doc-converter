namespace DocConverterApi.Helpers
{
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using Aspose.Words;
    using Aspose.Words.Drawing;
    using Aspose.Pdf.Facades;
    using System;

    public class AsposeWordsHelper
    {
        public AsposeWordsHelper()
        {
            // apply license
            Stream lStream = IOFunctions.ReadResourceStreamFromExecutingAssembly("Aspose.Words.lic");

            if (lStream != null)
            {
                License lic = new License();

                lic.SetLicense(lStream);
            }
        }

        public byte[] ConvertHtmlToPDF(string htmlContent, bool isLandscape = false, bool isA4 = false)
        {
            if (!string.IsNullOrEmpty(htmlContent))
            {
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                builder.InsertHtml(htmlContent);

                doc.Sections[0].PageSetup.Orientation = isLandscape ? Orientation.Landscape : Orientation.Portrait;
                doc.Sections[0].PageSetup.PaperSize = isA4 ? PaperSize.A4 : PaperSize.Letter;
                doc.UpdatePageLayout();

                MemoryStream outStream = new MemoryStream();
                doc.Save(outStream, SaveFormat.Pdf);

                outStream.Position = 0;
                return outStream.ToArray();
            }

            return null;
        }

        public byte[] ConvertToRTF(string fileLocation)
        {
            Document doc = new Document(fileLocation);

            MemoryStream docStream = new MemoryStream();

            doc.Save(docStream, SaveFormat.Rtf);
            return docStream.ToArray();
        }

        public byte[] ConvertToPDF(string fileLocation)
        {
            Document doc = new Document(fileLocation);

            MemoryStream docStream = new MemoryStream();

            doc.Save(docStream, SaveFormat.Pdf);
            return docStream.ToArray();
        }

        public byte[] ConvertToDOCX(string fileLocation)
        {
            Document doc = new Document(fileLocation);

            MemoryStream docStream = new MemoryStream();

            doc.Save(docStream, SaveFormat.Docx);
            return docStream.ToArray();
        }

        public byte[] ConvertToDOTX(string fileLocation)
        {
            Document doc = new Document(fileLocation);

            MemoryStream docStream = new MemoryStream();

            doc.Save(docStream, SaveFormat.Dotx);
            return docStream.ToArray();
        }

        public byte[] ConvertToDOC(string fileLocation)
        {
            Document doc = new Document(fileLocation);

            MemoryStream docStream = new MemoryStream();

            doc.Save(docStream, SaveFormat.Doc);
            return docStream.ToArray();
        }

        public bool CheckWordFileForDigitalSignatures(string fileLocation)
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(fileLocation);
            if (info.HasDigitalSignature)
                return true;
            else
                return false;
        }
    }
}
