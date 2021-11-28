using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace dot2doc
{
    public class DocumentGenerator
    {
        public static WordprocessingDocument CreateDocx(string filename)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(
                filename, WordprocessingDocumentType.Document))
            {
                // Add a new main document part and document dom. 
                package.AddMainDocumentPart().Document =
                    new Document(
                        new Body(
                            new Paragraph(
                                new Run(
                                    new Text("Hello World!")))));

                return package;
            }
        }
    }
}
