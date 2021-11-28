using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace dot2doc
{
    public class DocumentGenerator
    {
        public static WordprocessingDocument OpenDocument(string filename)
        {
            using (WordprocessingDocument wpDoc = WordprocessingDocument.Open(filename, true))
            {
                return wpDoc;
            }
        }

        public static WordprocessingDocument CreateDocx(string filename)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(
                filename, WordprocessingDocumentType.Document))
            {
                // Add a new main document part and document dom. 
                package.AddMainDocumentPart().Document = CreateEmptyDocPart();
                //package.MainDocumentPart.
     
                    //new Document(
                    //    new Body(
                    //        new Paragraph(
                    //            new Run(
                    //                new Text("Hello World!")))));


                return package;
            }
        }

        private static Header CreateEmptyDocHeader()
            => new Header();

        private static Document CreateEmptyDocPart()
            => new Document(
                new Body());

        public static void AddHeaderFromTo(WordprocessingDocument docFrom, WordprocessingDocument docTo)
        {
            MainDocumentPart mainPartTo = docTo.MainDocumentPart;

            // Delete the existing header part
            mainPartTo.DeleteParts(mainPartTo.HeaderParts);

            // Create a new header part
            HeaderPart headerPartTo =
                mainPartTo.AddNewPart<HeaderPart>();

            // Get Id of the headerPart
            string headerTo_Id = mainPartTo.GetIdOfPart(headerPartTo);

            // Feed target headerPart with source headerPart
            HeaderPart headerPartFrom =
                docFrom.MainDocumentPart.HeaderParts.FirstOrDefault();

            if (headerPartFrom != null)
            {
                headerPartTo.FeedData(headerPartFrom.GetStream());
            }

            // Get SectionProperties and Replace HeaderReference with new Id
            IEnumerable<SectionProperties> sectProps =
                mainPartTo.Document.Body.Elements<SectionProperties>();

            foreach (var sectProp in sectProps)
            {
                // Delete existing references to headers.
                sectProp.RemoveAllChildren<HeaderReference>();

                // Create the new header reference node.
                sectProp.PrependChild<HeaderReference>(new HeaderReference() { Id = headerTo_Id });
            }
        }

    }
}
