using System.Drawing;
using System;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;
using static System.Net.Mime.MediaTypeNames;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Metadata;
using System.Security.Cryptography.Xml;
using System.Xml.Linq;
using Microsoft.Extensions.Configuration;
using QRCoder;
using System.Linq;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;


namespace BMH
{
    //BookMarkHelper
    public class BMH
    {

        public static void UpdateTextBookmarks(string filePath, Dictionary<string, string> bookmarksContent)
        {
            bool inserted=false;
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                foreach (var entry in bookmarksContent)
                {
                    var bookmark= FindBookmark(wordDoc, entry.Key);
                    if (bookmark != null)
                    {
                        // Empty the content between the bookmarks
                        EmptyBookmark(wordDoc, bookmark.Value.bookmarkStart, bookmark.Value.bookmarkEnd);
                        // Insert the new content at the bookmark
                        inserted = InsertTextAtBookmark(wordDoc, bookmark.Value.bookmarkStart, entry.Value);
                    }
                    else {
                        Console.WriteLine("Bookmark "+ entry.Key+ "Not Found");
                    }
                    
                }
                if (inserted)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                }
                else {
                    //error ?
                    Console.WriteLine("Insertion Failed!");
                }
            }
        }


        public static (BookmarkStart? bookmarkStart, BookmarkEnd? bookmarkEnd)? FindBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            // Retrieve all BookmarkStart elements (from the main document, headers, footers, etc.)
            var allBookmarks = wordDoc.MainDocumentPart.Document
                .Descendants<BookmarkStart>()
                .Concat(wordDoc.MainDocumentPart.HeaderParts
                    .SelectMany(header => header.RootElement.Descendants<BookmarkStart>()))
                .Concat(wordDoc.MainDocumentPart.FooterParts
                    .SelectMany(footer => footer.RootElement.Descendants<BookmarkStart>()))
                .Concat(wordDoc.MainDocumentPart.FootnotesPart?.RootElement
                    .Descendants<BookmarkStart>() ?? Enumerable.Empty<BookmarkStart>())
                .Concat(wordDoc.MainDocumentPart.EndnotesPart?.RootElement
                    .Descendants<BookmarkStart>() ?? Enumerable.Empty<BookmarkStart>());

            // Find the first BookmarkStart element that matches the specified bookmark name
            var bookmarkStart = allBookmarks.FirstOrDefault(b => b.Name == bookmarkName);

            if (bookmarkStart != null)
            {
                // Call the method to find the BookmarkEnd using the bookmarkStart.Id
                var bookmarkEnd = FindBookmarkEnd(wordDoc, bookmarkStart.Id);

                if (bookmarkEnd != null)
                {
                    // Return both BookmarkStart and BookmarkEnd as a tuple if both are found
                    return (bookmarkStart, bookmarkEnd);
                }
            }

            // Return null if no matching BookmarkStart or BookmarkEnd was found
            return null;
        }

        public static BookmarkEnd? FindBookmarkEnd(WordprocessingDocument wordDoc, string bookmarkId)
        {
            // Retrieve all BookmarkEnd elements and find the one that matches the given Id
            var bookmarkEnd = wordDoc.MainDocumentPart.Document.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => b.Id == bookmarkId);

            return bookmarkEnd;
        }


        public static void EmptyBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, BookmarkEnd bookmarkEnd)
        {
            // List to hold elements to remove
            var elementsToRemove = new List<OpenXmlElement>();

            // Start from the sibling after the BookmarkStart element and iterate until the BookmarkEnd
            for (var currentElement = bookmarkStart.NextSibling(); currentElement != null && currentElement != bookmarkEnd; currentElement = currentElement.NextSibling())
            {
                // Collect all elements in the range to be removed
                elementsToRemove.Add(currentElement);

                // Recursively collect child elements if needed (for nested elements)
                CollectNestedElements(currentElement, elementsToRemove);
            }

            // Remove all collected elements
            foreach (var element in elementsToRemove)
            {
                // Safety check in case something is unexpectedly null
                if (element != null)
                {
                    element.Remove();
                }
            }
        }

        // Helper method to remove child elements recursively in EmptyBookmark method
        private static void CollectNestedElements(OpenXmlElement element, List<OpenXmlElement> elementsToRemove)
        {
            if (element == null) return; // Safety check for null

            // Loop through all child elements and add them to the list to be removed
            foreach (var child in element.Elements())
            {
                if (child == null) continue; // Safety check for null child

                elementsToRemove.Add(child);
                // Recursively collect child elements if they have their own children
                CollectNestedElements(child, elementsToRemove);
            }
        }


        //most simple
        //public static bool InsertTextAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string content) {
        //    var newRun = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(content));
        //    bookmarkStart.Parent.InsertAfter(newRun, bookmarkStart);
        //    return true;
        //}



        public static bool InsertTextAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        {

            try
            {
                // Debug: Log the bookmark name and new text
                Console.WriteLine($"Attempting to insert text at bookmark. Bookmark name: {bookmarkStart.Name}, New text: {newText}");

                // Find the parent element of the bookmark
                var parentElement = bookmarkStart.Parent;

                if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
                {
                    // Debug: Log that the parent is a paragraph
                    Console.WriteLine("Parent element is a Paragraph.");

                    // Find the first Run after the bookmark to clone its style (if needed)
                    var existingRun = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();

                    if (existingRun != null)
                    {
                        Console.WriteLine("Found an existing Run to clone style.");
                    }
                    else
                    {
                        Console.WriteLine("No existing Run found. Creating a new one.");
                    }

                    var run = existingRun != null
                        ? (DocumentFormat.OpenXml.Wordprocessing.Run)existingRun.CloneNode(true)
                        : new DocumentFormat.OpenXml.Wordprocessing.Run();

                    // Debug: Log that we're removing existing text from the cloned Run
                    Console.WriteLine("Removing existing text from the Run.");
                    run.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.Text>();

                    // Debug: Log that we're appending the new text
                    Console.WriteLine("Appending new text to the Run.");
                    run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

                    // Debug: Log that we're inserting the Run after the bookmark
                    Console.WriteLine("Inserting the Run after the BookmarkStart.");
                    paragraph.InsertAfter(run, bookmarkStart);
                }
                else if (parentElement != null)
                {
                    // Debug: Log that the parent is not a paragraph
                    Console.WriteLine($"Parent element is not a Paragraph. Parent type: {parentElement.GetType().Name}");

                    // Handle non-paragraph parents (e.g., tables)
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

                    // Debug: Log that we're inserting the Run after the bookmark
                    Console.WriteLine("Inserting the Run after the BookmarkStart.");
                    parentElement.InsertAfter(run, bookmarkStart);
                }
                else
                {
                    // Debug: Log that the parent element is null
                    Console.WriteLine("Parent element is null. Unable to insert text.");
                    return false; // Parent element couldn't be determined
                }

                // Debug: Log success
                Console.WriteLine("Text insertion successful.");
                return true; // Insertion successful
            }
            catch (Exception ex)
            {
                // Debug: Log the exception
                Console.WriteLine($"Error during text insertion: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                return false; // Insertion failed due to an exception
            }
        }













        //refactored simple inserttextatbookmark
        //public static void InsertTextAtBookmark2(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        //{
        //    // Find the parent element of the bookmark
        //    var parentElement = bookmarkStart.Parent;

        //    if (parentElement is DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph)
        //    {
        //        // Insert a new Run containing the text directly into the existing Paragraph
        //        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));
        //        paragraph.InsertAfter(run, bookmarkStart);
        //    }
        //    else if (parentElement != null)
        //    {
        //        // For non-paragraph parents (like tables), add the new content in the parent
        //        var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));
        //        parentElement.InsertAfter(run, bookmarkStart);
        //    }
        //    else
        //    {
        //        throw new InvalidOperationException("The bookmark's parent element could not be determined.");
        //    }
        //}

        //function made inserttextbookmark
        //public static void InsertTextAtBookmark(WordprocessingDocument wordDoc, BookmarkStart bookmarkStart, string newText)
        //{
        //    // Create a new run with the new text content (use Wordprocessing namespace)
        //    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(newText));

        //    // Create a new paragraph to hold the new run (use Wordprocessing namespace)
        //    var paragraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run);

        //    // Find the parent element (usually a paragraph) of the BookmarkStart
        //    var parentElement = bookmarkStart.Parent as OpenXmlElement;

        //    if (parentElement != null)
        //    {
        //        // Insert the new content at the position where the bookmark is located
        //        parentElement.InsertAfter(paragraph, bookmarkStart);
        //    }

        //}


    }
}
