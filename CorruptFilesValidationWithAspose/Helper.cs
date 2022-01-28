using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using Aspose.Words;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace FilesWithAspose
{
    class Helper
    {
        public static void Invoker()
        {
            Helper helper = new Helper();
            helper.ClickEventExecuted();
        }

        private void ClickEventExecuted()
        {
            //Test_ASPOSE_PDF();
            //TestBugWith_ID866411();
            //TestBugWith_ID868769();
            //ExtractTextFromPDF();            

            //=======================================================================================
            var filePath = GetFilePath();
            byte[] payLoad = System.IO.File.ReadAllBytes(filePath);
            var validationList = GetFileValidationWithAsposeLibrary(filePath, payLoad);
            //=======================================================================================

            //CheckWordFileCorruptedException(GetFilePath(), payLoad);

            //CheckWordFileCorruptedException_v2(GetFilePath(), payLoad);
        }

        private string GetFilePath()
        {
            string filepath = string.Empty;

            #region Sample Payloads
            filepath = @"../../Payload/MyAgreement Document.docx"; //MyAgreement Document
            //filepath = @"../../Payload/MyFirstAssociation Document.docx"; //MyAgreement Document
            //filepath = @"../../Payload/MySecondAssociation Document.docx"; //MyAgreement Document
            #endregion

            //filepath = @"C:\Users\c-rajesh.thomas\Downloads\Amit P Corrupted File\COG-AOG-SOW-2021-01799_Project+Saturn (1).docx"; //AP Corrupted File
            //filepath = @"C:\Users\c-rajesh.thomas\Downloads\Amit P Corrupted File\1 + -+Requisition_Data_Export_2021 - 11 - 30(1).pdf";  //AP Corrupted File

            // ==================================================================================================================================
            // ==================================================================================================================================

            //filepath = @"C:\Users\c-rajesh.thomas\Downloads\1000332329.pdf"; //RW -- v1
            // filepath = @"C:\Users\c-rajesh.thomas\Downloads\AsposeWordCorrupted\1000350128.docx"; //RW -- v2

            // ==================================================================================================================================
            // ==================================================================================================================================

            //string RootFolder_Bug868769 = @"C:\TFS\RT_AllData\Desktop\RT\ICERTIS\Bugs Lists\Bug-868769 CTS Corrupted PDF File\";
            //filepath = RootFolder_Bug868769 + "Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed.Pdf";
            //filepath = RootFolder_Bug868769 + "ICMMSASellsideAmendment_cc1d564b-0cfd-4e6b-8113-f8e93068fc62_0.Pdf";

            //filepath = RootFolder_Bug868769 + "Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed_Page3_WithLink.Pdf";
            //filepath = RootFolder_Bug868769 + "CorruptFile.Pdf";
            //filepath = RootFolder_Bug868769 + "Test.Pdf";

            // ==================================================================================================================================
            // ==================================================================================================================================

            //string RootFolder_Bug946617 = @"C:\TFS\RT_AllData\Desktop\RT\ICERTIS\Bugs Lists\Bug-946617_BCG_PDF_Signed_Copy\";
            //filepath = RootFolder_Bug946617 + "Signed_21.03.22_ICA_Dolores_Lee_Hong_Yeen_ASv.3_signed_signed.Pdf";
            //filepath = RootFolder_Bug946617 + "BasicTestFile.pdf";
            //filepath = RootFolder_Bug946617 + @"SimplyCopied\Signed_21.03.22_ICA_Dolores_Lee_Hong_Yeen_ASv.3_signed_signed - Copy.pdf";
            //filepath = RootFolder_Bug946617 + @"1-7 Pages\Signed_21.03.22_ICA_Dolores_Lee_Hong_Yeen_ASv.3_signed_signed_1-7Pages.Pdf";

            // ==================================================================================================================================
            // ==================================================================================================================================
            return filepath;
        }

        #region Amit and CTS File Corrupted Issue POC

        static string Rewriter(Uri partUri, string id, string uri) => $"http://unknown?id={id}";
        public static void CheckWordFileCorruptedException_v2(string filePath, byte[] file)
        {
            try
            {
                var fileBytes = file;
                // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
                // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
                // --- Aspose.Words.Document doc = new Aspose.Words.Document(filepath);

                using (var fileStream = new MemoryStream())
                {
                    //WordprocessingDocument.Open(fileStream, true);

                    //Aspose.Words.LoadOptions loadOptions = new Aspose.Words.LoadOptions();
                    //loadOptions.PreserveIncludePictureField = true;
                    //using (MemoryStream docxStream = new MemoryStream(fileBytes))
                    //{
                    //    Aspose.Words.Document doc = new Aspose.Words.Document(docxStream, loadOptions);
                    //    doc.Save(@"C:\Users\c-rajesh.thomas\Downloads\AsposeWordCorrupted\1000350128_CreatedByWpfApp1_" + Guid.NewGuid() + ".docx");
                    //}

                    // Approach: 1 for OpenSettings with default handler
                    // This works with the latest package of Open XML and wont work for 2.5.0 version
                    //var openSettings = new OpenSettings()
                    //{
                    //    RelationshipErrorHandlerFactory = RelationshipErrorHandler.CreateRewriterFactory(Rewriter)
                    //};

                    // Approach: 2 for OpenSettings with custom handler
                    // This works with the latest package of Open XML and wont work for 2.5.0 version
                    var openSettings = new OpenSettings()
                    {
                        RelationshipErrorHandlerFactory = package =>
                        {
                            return new UriRelationshipErrorHandler();
                        }
                    };

                    WordprocessingDocument wDoc;
                    try
                    {
                        using (wDoc = WordprocessingDocument.Open(filePath, true, openSettings))
                        {
                            ProcessDocument(wDoc);
                        }
                    }
                    catch (OpenXmlPackageException e)
                    {
                        if (e.ToString().Contains("Invalid Hyperlink"))
                        {
                            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                            {
                                UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                            }
                            using (wDoc = WordprocessingDocument.Open(filePath, true))
                            {
                                var resultCount = ProcessDocument(wDoc);
                            }
                        }
                    }
                }

            }
            catch (FileCorruptedException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static Uri FixUri(string brokenUri)
        {
            return new Uri("http://broken-link/");
        }


        private static int ProcessDocument(WordprocessingDocument wDoc)
        {
            var elementCount = wDoc.MainDocumentPart.Document.Descendants().Count();
            //Console.WriteLine(elementCount);
            return elementCount;
        }



        #endregion

        public static void CheckWordFileCorruptedException(string filePath, byte[] file)
        {
            try
            {
                var fileBytes = file;
                // If we get an "Unreadable content" error message when trying to open a document using Microsoft Word,
                // chances are that we will get an exception thrown when trying to load that document using Aspose.Words.
                // --- Aspose.Words.Document doc = new Aspose.Words.Document(filepath);

                using (var fileStream = new MemoryStream())
                {
                    WordprocessingDocument.Open(fileStream, true);

                    //Aspose.Words.LoadOptions loadOptions = new Aspose.Words.LoadOptions();
                    //loadOptions.PreserveIncludePictureField = true;
                    //using (MemoryStream docxStream = new MemoryStream(fileBytes))
                    //{
                    //    Aspose.Words.Document doc = new Aspose.Words.Document(docxStream, loadOptions);
                    //    doc.Save(@"C:\Users\c-rajesh.thomas\Downloads\AsposeWordCorrupted\1000350128_CreatedByWpfApp1_" + Guid.NewGuid() + ".docx");
                    //}
                }

            }
            catch (FileCorruptedException e)
            {
                Console.WriteLine(e.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// GetFileValidation in ICI
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="file"></param>
        /// <returns></returns>
        public static List<string> GetFileValidationWithAsposeLibrary(string filePath, byte[] file)
        {
            List<string> validationMessage = new List<string>();
            var fileExt = System.IO.Path.GetExtension(filePath).ToUpperInvariant();
            var fileBytes = file;

            if ((!string.IsNullOrEmpty(filePath) && fileBytes == null) || (fileBytes != null))
            {
                try
                {
                    switch (fileExt)
                    {
                        case ".DOCX":
                        case ".DOC":
                            using (var fileStream = new MemoryStream())
                            {
                                // This validates the DOCX file using Aspose Library
                                Aspose.Words.LoadOptions loadOptions = new Aspose.Words.LoadOptions();
                                loadOptions.PreserveIncludePictureField = true;
                                using (MemoryStream docxStream = new MemoryStream(fileBytes))
                                {
                                    new Aspose.Words.Document(docxStream, loadOptions);
                                }

                                // This validates the DOCX file using OpenXML Library
                                fileStream.Write(fileBytes, 0, fileBytes.Length);
                                WordprocessingDocument.Open(fileStream, true);
                            }
                            break;

                        case ".PDF":
                            using (MemoryStream memoryStream = new MemoryStream(fileBytes))
                            {
                                List<string> ErrorMessages = new List<string>();

                                PdfFileInfo info = new PdfFileInfo(filePath);
                                if (!info.IsPdfFile)
                                {
                                    ErrorMessages.Add("Invalid PDF File");
                                }

                                //Document doc = new Document();
                                //doc.EnableSignatureSanitization = false;
                                //Document mDoc = new Document(new FileStream("Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed - Copy.Pdf", FileMode.Open));
                                //doc.Pages.Add(mDoc.Pages);

                                var pdfFileEditor = new Aspose.Pdf.Facades.PdfFileEditor();
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorrupted; //Default

                                //Start :: For Bug 868769
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.StopWithError; // This Works
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorrupted; // This Doesnt Work ==> ERROR
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorruptedObjects; // This Works - Checked-In
                                //End :: For Bug 868769

                                //Start :: For Bug 946617
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.StopWithError; // Only This Works
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorrupted; // This Doesnt Work ==> ERROR
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorruptedObjects; // This Doesnt Work ==> ERROR                                //End :: For Bug 946617

                                //Start :: PS Team
                                pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.StopWithError; // This Works
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorrupted; // This Works
                                //pdfFileEditor.CorruptedFileAction = Aspose.Pdf.Facades.PdfFileEditor.ConcatenateCorruptedFileAction.ConcatenateIgnoringCorruptedObjects; // This Works
                                //End :: PS Team


                                MemoryStream outputStream = new MemoryStream();
                                pdfFileEditor.Concatenate(new Stream[] { memoryStream }, outputStream);

                                if (pdfFileEditor.CorruptedItems.Length > 0)
                                {
                                    foreach (var a in pdfFileEditor.CorruptedItems)
                                    {
                                        ErrorMessages.Add(a.Exception.Message);
                                        validationMessage.Add(a.Exception.Message);
                                    }

                                    throw new Exception("CORRUPTED FILE");
                                }
                            }
                            break;
                    }
                }
                catch (FileFormatException)
                {
                }
            }
            return validationMessage;
        }

        // ====================================================================================================

        public static void Test_ASPOSE_PDF()
        {
            var fileName = "TestDoc.pdf";
            using (var pdfDocument = new Aspose.Pdf.Document(fileName))
            {
                Console.WriteLine($"Pages {pdfDocument.Pages.Count}");
            }
        }
        public static void TestBugWith_ID866411()
        {
            //For Bug-866411 from 7.7
            string path = "C:/Users/c-rajesh.thomas/Downloads/Bug-866411/Nexidia_Nice - SOW 4 (renewal of SOW 1&2)_4.Pdf";
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(path);
        }

        public static void TestBugWith_ID868769()
        {
            //For Bug-868769 from 7.11
            string path = "C:/Users/c-rajesh.thomas/Downloads/Bug-868769/ICMMSASellsideAmendment_cc1d564b-0cfd-4e6b-8113-f8e93068fc62_0.Pdf";
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(path);

            //For Bug-868769 from 7.11
            string path1 = "C:/Users/c-rajesh.thomas/Downloads/Bug-868769/Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed.pdf";
            Aspose.Pdf.Document pdfDoc1 = new Aspose.Pdf.Document(path1);
        }
        public static void ExtractTextFromPDF()
        {
            string MyDirectory = "";

            // Open PDF document
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDirectory + "Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed.Pdf");

            // Document pdfDocument = new Document(@"C:/Users/c-rajesh.thomas/Downloads/Bug-868769/Sixth Amendment to ADM Managed Services MSA between Mattel and Cognizant - signed_Page1.pdf");
            //Document pdfDocument = new Document(@"C:/Users/c-rajesh.thomas/Downloads/Bug-868769/TestDoc.pdf");

            // Create TextAbsorber object to extract text
            TextAbsorber textAbsorber = new TextAbsorber();

            // Accept the absorber for all pages
            pdfDocument.Pages.Accept(textAbsorber);

            // Get the extracted text
            string extractedText = textAbsorber.Text;

            // Create a writer and open the file
            TextWriter tw = new StreamWriter(MyDirectory + "extracted-text.txt");

            // Write a line of text to the file
            tw.WriteLine(extractedText);

            // Close the stream
            tw.Close();
        }


    }


    public static class UriFixer
    {
        public static void FixInvalidUri(Stream fs, Func<string, Uri> invalidUriHandler)
        {
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (ZipArchive za = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                foreach (var entry in za.Entries.ToList())
                {
                    if (!entry.Name.EndsWith(".rels"))
                        continue;
                    bool replaceEntry = false;
                    XDocument entryXDoc = null;
                    using (var entryStream = entry.Open())
                    {
                        try
                        {
                            entryXDoc = XDocument.Load(entryStream);
                            if (entryXDoc.Root != null && entryXDoc.Root.Name.Namespace == relNs)
                            {
                                var urisToCheck = entryXDoc
                                    .Descendants(relNs + "Relationship")
                                    .Where(r => r.Attribute("TargetMode") != null && (string)r.Attribute("TargetMode") == "External");
                                foreach (var rel in urisToCheck)
                                {
                                    var target = (string)rel.Attribute("Target");
                                    if (target != null)
                                    {
                                        try
                                        {
                                            Uri uri = new Uri(target);
                                        }
                                        catch (UriFormatException)
                                        {
                                            Uri newUri = invalidUriHandler(target);
                                            rel.Attribute("Target").Value = newUri.ToString();
                                            replaceEntry = true;
                                        }
                                    }
                                }
                            }
                        }
                        catch (XmlException)
                        {
                            continue;
                        }
                    }
                    if (replaceEntry)
                    {
                        var fullName = entry.FullName;
                        entry.Delete();
                        var newEntry = za.CreateEntry(fullName);
                        using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                        using (XmlWriter xmlWriter = XmlWriter.Create(writer))
                        {
                            entryXDoc.WriteTo(xmlWriter);
                        }
                    }
                }
            }
        }
    }

    public class UriRelationshipErrorHandler : RelationshipErrorHandler
    {
        public override string Rewrite(Uri partUri, string id, string uri)
        {
            return "http://link-invalido";
        }
    }

}
