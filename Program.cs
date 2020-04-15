using System;
using System.IO;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace toc_generator
{
    class Program
    {

        // patterns of section titles
        private static readonly Regex BEGIN_SECTION = new Regex(@"(^\d\.\d.[^\d])|(^[A-Z]\d.[^\d])|(^この章)|(^目次)|(^目標)", RegexOptions.Compiled);
        // patterns of subsection titles
        private static readonly Regex BEGIN_SUB_SECTION = new Regex(@"(^\d\.\d\.\d)|(^[A-Z]\d\.\d)", RegexOptions.Compiled);
        // patterns of chapter titles
        private static readonly Regex BEGIN_CHAPTER = new Regex(@"(^第.章)|(^\d\.[^\d])|(^§)", RegexOptions.Compiled);
        // font size
        private const int pt = 11;
        // global counter to maintain 
        private static int paragraphCounter = 0;
        /// <summary>
        /// Generates TOC for each PowerPoint document as a Word document
        /// </summary>
        /// <param name="rootPath">The root path that contains all PowerPoint documents to be processed</param>
        static void GenerateTOC(string rootPath)
        {
            var objPowerPoint = new PowerPoint.Application(){
                DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone
            };
            var objWord = new Word.Application(){
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            };
            try
            {
                foreach (var file in Directory.GetFiles(rootPath, "*.ppt?", SearchOption.AllDirectories))
                {
                    if (file.Contains("reference")) continue;
                    var inputPath = Path.GetFullPath(file);
                    var dir_name = Path.GetDirectoryName(inputPath);
                    var file_name = Path.GetFileNameWithoutExtension(inputPath);
                    var outputPath = Path.Combine(dir_name, file_name + ".toc.docx");
                    bool listingSubSections = false;
                    // skipping documents that have been already converted
                    if (File.Exists(outputPath) && File.GetLastWriteTime(outputPath) >= File.GetLastWriteTime(inputPath))
                    {
                        Console.WriteLine($"skipping {inputPath}");
                        continue;
                    }
                    else
                    {
                        Console.WriteLine($"Extracting toc for {inputPath}");
                        var pptDoc = objPowerPoint.Presentations.Open(inputPath);
                        var wordDoc = objWord.Documents.Add(Visible: false);
                        for (int i = 1; i <= pptDoc.Slides.Count; ++i)
                        {
                            var slide = pptDoc.Slides[i];
                            var title = slide.Shapes.Title.TextFrame.TextRange.Text
                                        .Replace("\v", "")
                                        .Replace("\t", "")
                                        .Replace("\r\n", "");
                            var pageNum = slide.SlideNumber;
                            if (title == string.Empty)
                            {
                                continue;
                            }
                            else if (BEGIN_SUB_SECTION.IsMatch(title))
                            {
                                Console.WriteLine($"subsection: {title}");
                                NewParagraphForSubSections(wordDoc, $"{title}\t{pageNum}");
                                listingSubSections = true;
                            }
                            else if (BEGIN_SECTION.IsMatch(title))
                            {
                                Console.WriteLine($"section: {title}");
                                NewParagraphForSections(wordDoc, $"{title}\t{pageNum}");
                                listingSubSections = true;
                            }
                            else if (BEGIN_CHAPTER.IsMatch(title))
                            {
                                Console.WriteLine($"chapter: {title}");
                                if (i != 1) NewParagraph(wordDoc, "");
                                NewParagraphForChapters(wordDoc, $"{title}\t{pageNum}");
                                listingSubSections = true;
                            }
                            else if (title.Contains("付録") || title.Contains("Appendix"))
                            {
                                Console.WriteLine($"Chapter: {title}");
                                NewParagraphForChapters(wordDoc, $"{title}\t{pageNum}");
                                listingSubSections = true;
                            }
                            else if (listingSubSections)
                            {
                                Console.WriteLine($"subsection: {title}");
                                NewParagraphForSubSections(wordDoc, $"{title}\t{pageNum}");
                            }
                            else
                            {
                                Console.WriteLine($"section: {title}");
                                NewParagraphForSections(wordDoc, $"{title}\t{pageNum}");
                            }
                        }
                        AddHeader(wordDoc, file_name);
                        AddFooter(wordDoc, file_name);
                        wordDoc.SaveAs2(outputPath);
                        try
                        {
                            wordDoc.Close(SaveChanges:false);
                            pptDoc.Close();
                        }
                        catch(Exception e)
                        {
                            Console.Error.WriteLine($"An exception has been thrown @ closing the PowerPoint and Word documents: {e.Message}");
                        }
                        Console.WriteLine($"Completed: {outputPath}.");
                    }
                }
            }
            finally
            {
                objPowerPoint.Quit();
                objWord.Quit();
            }
        }
        /// <summary>
        /// Obtains the footer and header names for the given file. 
        /// </summary>
        /// <param name="fileName">The file name of the PowerPoint document</param>
        /// <returns> 
        /// <list>
        /// <item>"本編目次" if the fileName contains "main" (for example "05_main.pptx")</item>
        /// <item> "付録目次" if the fileName contains "appendix" (for example "06_appendix.pptx") </item>
        /// <item> "目次" otherwise </item>
        /// </list>
        /// </returns>
        private static string FooterHeaderName(string fileName)
        {
            if (fileName.Contains("main")) return "本編目次";
            else if (fileName.Contains("appendix")) return "付録目次";
            else return "目次";
        }

        /// <summary>
        /// Adds headers to Word documents
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="fileName">The file name of the PowerPoint document</param>
        private static void AddHeader(Word.Document wordDoc, string fileName)
        {
            foreach (Word.Section section in wordDoc.Sections)
            {
                foreach (Word.HeaderFooter header in section.Headers)
                {
                    header.Range.Paragraphs.Add();
                    var targetRange = header.Range.Paragraphs.Add().Range;
                    header.Range.Borders.Enable = 1;
                    header.Range.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                    header.Range.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleThinThickSmallGap;
                    header.Range.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    header.Range.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    header.Range.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    //header.Range.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth025pt;
                    targetRange.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleThinThickSmallGap;
                    targetRange.Text += $"【{FooterHeaderName(fileName)}】";
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }
            }
        }

        /// <summary>
        /// Adds footers to Word documents
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="fileName">The file name of the PowerPoint document</param>
        private static void AddFooter(Word.Document wordDoc, string fileName)
        {
            foreach (Word.Section section in wordDoc.Sections)
            {
                foreach (Word.HeaderFooter footer in section.Footers)
                {
                    footer.Range.Borders.Enable = 1;
                    footer.Range.Fields.Add(footer.Range, Word.WdFieldType.wdFieldPage);
                    footer.Range.InsertBefore($"【{FooterHeaderName(fileName)}－");
                    footer.Range.InsertAfter("】");
                    footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
        }

        /// <summary>
        /// Creates chapter entries
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="str">The title of the chapter</param>
        /// <returns></returns>
        private static Word.Paragraph NewParagraphForChapters(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 1.35F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 1;
            return p;
        }
        /// <summary>
        /// Creates section entries
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="str">The title of the section</param>
        /// <returns></returns>
        private static Word.Paragraph NewParagraphForSections(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 2.02F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 0;
            return p;
        }
        /// <summary>
        /// Creates subsection entries
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="str">The title of the subsection</param>
        /// <returns></returns>
        private static Word.Paragraph NewParagraphForSubSections(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 3.37F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 0;
            return p;
        }
        /// <summary>
        /// Creates a paragraph object
        /// </summary>
        /// <param name="wordDoc">The Word document</param>
        /// <param name="str">The content of the paragraph</param>
        /// <returns></returns>
        private static Word.Paragraph NewParagraph(Word.Document wordDoc, string str)
        {
            var wordParagraph = wordDoc.Content.Paragraphs.Add();
            if (paragraphCounter++ == 0)
                wordParagraph.Range.Text = str;
            else
                wordParagraph.Range.Text += str;
            wordParagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            wordParagraph.Range.ParagraphFormat.TabStops.ClearAll();
            wordParagraph.Range.ParagraphFormat.TabStops.Add(40 * pt, Alignment: Word.WdTabAlignment.wdAlignTabRight, Leader: Word.WdTabLeader.wdTabLeaderDots);
            return wordParagraph;
        }
        static void Main(string[] args)
        {
            GenerateTOC(args[0]);
        }
    }
}
