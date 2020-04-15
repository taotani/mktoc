using System;
using System.IO;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace toc_generator
{
    class Program
    {
        private static readonly Regex BEGIN_SECTION = new Regex(@"(^\d\.\d.[^\d])|(^[A-Z]\d.[^\d])|(^この章)|(^目次)|(^目標)", RegexOptions.Compiled);
        private static readonly Regex BEGIN_SUB_SECTION = new Regex(@"(^\d\.\d\.\d)|(^[A-Z]\d\.\d)", RegexOptions.Compiled);
        private static readonly Regex BEGIN_CHAPTER = new Regex(@"(^第.章)|(^\d\.[^\d])|(^§)", RegexOptions.Compiled);
        private const int pt = 11;
        private static int paragraphCounter = 0;
        static void GenerateTOC(string rootPath)
        {
            var objPowerPoint = new PowerPoint.Application();
            var objWord = new Word.Application();
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
                        //wordDoc.Close();
                        //pptDoc.Close();
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
        private static string FooterHeaderName(string fileName)
        {
            if (fileName.Contains("main")) return "本編目次";
            else if (fileName.Contains("appendix")) return "付録目次";
            else return "目次";
        }
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
        private static Word.Paragraph NewParagraphForChapters(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 1.35F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 1;
            return p;
        }
        private static Word.Paragraph NewParagraphForSections(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 2.02F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 0;
            return p;
        }
        private static Word.Paragraph NewParagraphForSubSections(Word.Document wordDoc, string str)
        {
            var p = NewParagraph(wordDoc, str);
            p.Range.ParagraphFormat.LeftIndent = 3.37F * pt;
            p.Range.ParagraphFormat.RightIndent = 1F * pt;
            p.Range.Bold = 0;
            return p;
        }
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


// Sub createIndex()
//     'ファイルのオープン
//     Open "TableOfContents.txt" For Output As #1
//         'ヘッダの印字
//         Print #1, "タイトル"; vbTab; "ページ"
//         For Each Slide In Application.ActivePresentation.Slides
//             'スライドタイトルの取得
//             Title = Slide.Shapes.Title.TextFrame.TextRange.Text
//             '非表示文字の削除
//             Title = Replace(Title, vbVerticalTab, "") '垂直タブ
//             Title = Replace(Title, vbTab, "")         '水平タブ
//             Title = Replace(Title, vbNewLine, "")     '改行文字
//             'ページ番号の取得
//             pageNum = Slide.SlideNumber
//             'ファイルへ出力
//             Print #1, Title; vbTab; pageNum
//         Next
//     'ファイルのクローズ
//     Close #1
// End Sub