// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtSaveOptions : ApiExampleBase
    {
        [Test]
        public void PageBreaks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.ForcePageBreaks
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document(MyDir + "SaveOptions.PageBreaks.docx");
            TxtSaveOptions saveOptions = new TxtSaveOptions {ForcePageBreaks = false};

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PageBreaks.False.txt", saveOptions);
            //ExEnd
            Document docFalse = new Document(MyDir + @"\Artifacts\SaveOptions.PageBreaks.False.txt");
            Assert.AreEqual(
                "Some text before page break\r\rJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\rqwdqwdqwdqwdqwdqwdqwqwd\r\f",
                docFalse.GetText());

            saveOptions.ForcePageBreaks = true;
            doc.Save(MyDir + @"\Artifacts\SaveOptions.PageBreaks.True.txt", saveOptions);

            Document docTrue = new Document(MyDir + @"\Artifacts\SaveOptions.PageBreaks.True.txt");
            Assert.AreEqual(
                "Some text before page break\r\f\r\fJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\f\r\fqwdqwdqwdqwdqwdqwdqwqwd\r\f",
                docTrue.GetText());
        }

        [Test]
        public void AddBidiMarks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document(MyDir + "Document.docx");
            // In Aspose.Words by default this option is set to true unlike Word
            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = false };

            doc.Save(MyDir + @"\Artifacts\AddBidiMarks.txt", saveOptions);
            //ExEnd
        }
    }
}