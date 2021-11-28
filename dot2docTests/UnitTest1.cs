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
    }
}
