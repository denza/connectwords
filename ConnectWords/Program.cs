using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static XNamespace r =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    static string altChunkId = "AltChunkId";

    static void Main(string[] args)
    {
        var docxs = new List<byte[]>();
        docxs.Add(File.ReadAllBytes("myDoc1.docx"));
        docxs.Add(File.ReadAllBytes("myDoc2.docx"));
        docxs.Add(File.ReadAllBytes("myDoc3.docx"));

        var newDoc = "mergeDoc.docx";

        GenerateDocument(newDoc,docxs);
    }

    /// <summary>
    /// Gets a memory stream with prefilled data, and sets it to be read
    /// </summary>
    /// <param name="fileData">Data preload in a stream</param>
    /// <returns>Memory stream with prefilled data</returns>
    private static Stream GetStream(byte[] fileData)
    {

        var ms = new MemoryStream();

        ms.Write(fileData, 0, fileData.Length);
        ms.Position = 0;

        return ms;
    }

    /// <summary>
    /// Generats a new word document
    /// </summary>
    /// <param name="path">Path where a new document will be created</param>
    private static void GenerateNewDocx(string path)
    {
        using (WordprocessingDocument mainDoc = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = mainDoc.MainDocumentPart == null ? mainDoc.AddMainDocumentPart() : mainDoc.MainDocumentPart;
            if (mainPart.Document == null)
            {
                mainPart.Document = new Document(new Body(new Paragraph()));
            }

            mainDoc.Close();
        }
    }

    /// <summary>
    /// In a document body finds free id for altChunk
    /// </summary>
    /// <param name="xdoc"></param>
    /// <returns>Returns next id for altChunk</returns>
    private static string findNextIdForAltChunk(XDocument xdoc)
    {
        int nextId = 0;

        //Loop until you find a free number
        while (true)
        {
            var idUsed = xdoc.Root
              .Element(w + "body")
              .Elements(w + "altChunk").Any(el =>
                  {
                      XAttribute atr = el.Attributes().FirstOrDefault(x => x.Name.LocalName.Contains("id"));
                      if (atr == null)
                          return false;

                      if (atr.Value == (altChunkId + nextId))
                          return true;

                      return false;
                  });

            if (idUsed)
                nextId++;
            else
                break;
        }

        return altChunkId + nextId;
    }

    /// <summary>
    /// Merges docxs in a new docx
    /// </summary>
    /// <param name="path">A new document destination</param>
    /// <param name="documents">Documents to be merged</param>
    private static void GenerateDocument(string path, List<byte[]> documents)
    {
        GenerateNewDocx(path);

        using (WordprocessingDocument mainDoc = WordprocessingDocument.Open(path, true))
        {
            MainDocumentPart mainPart = mainDoc.MainDocumentPart;
            XDocument mainDocumentXDoc = GetXDocument(mainDoc);

            foreach (var item in documents)
            {
                var currentAltChunkId = findNextIdForAltChunk(mainDocumentXDoc);
                AlternativeFormatImportPart chunk = mainPart.AddAlternativeFormatImportPart(
               "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
               currentAltChunkId);

                chunk.FeedData(GetStream(item));

                XElement altChunk = new XElement(w + "altChunk", new XAttribute(r + "id", currentAltChunkId));

                mainDocumentXDoc.Root
                                .Element(w + "body")
                                .Elements()
                                .Last()
                                .AddAfterSelf(altChunk);
            }

            SaveXDocument(mainDoc, mainDocumentXDoc);
        }
    }

    /// <summary>
    /// Saves xml  as docx
    /// </summary>
    /// <param name="myDoc"></param>
    /// <param name="mainDocumentXDoc"></param>
    private static void SaveXDocument(WordprocessingDocument myDoc, XDocument mainDocumentXDoc)
    {
        // Serialize the XDocument back into the part
        using (Stream str = myDoc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            using (XmlWriter xw = XmlWriter.Create(str))
            {
                mainDocumentXDoc.Save(xw);
            }
        }
    }

    /// <summary>
    /// Gets xml from docx
    /// </summary>
    /// <param name="myDoc"></param>
    /// <returns></returns>
    private static XDocument GetXDocument(WordprocessingDocument myDoc)
    {
        // Load the main document part into an XDocument
        XDocument mainDocumentXDoc;
        using (Stream str = myDoc.MainDocumentPart.GetStream())
        {
            using (XmlReader xr = XmlReader.Create(str))
            {
                mainDocumentXDoc = XDocument.Load(xr);
            }
        }

        return mainDocumentXDoc;
    }
}