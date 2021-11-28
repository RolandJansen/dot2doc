using DocumentFormat.OpenXml.Packaging;
using dot2doc;
using System;
using Xunit;

namespace dot2docTests
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            DocumentGenerator.CreateDocx("TestDoc.docx");
        }

        [Fact]
        public void Test2()
        {
            WordprocessingDocument template =
                DocumentGenerator.OpenDocument(@"C:\Users\JansenR\Documents\dot2doc_test\header_copy_test.dotx");

            WordprocessingDocument output = DocumentGenerator.CreateDocx("header_copy_output.docx");
            //DocumentGenerator.AddHeaderFromTo(template, output);

            output.Save();
        }
    }
}
