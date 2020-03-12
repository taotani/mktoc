using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace index_generator
{
    class Program
    {
        static void GenerateIndex(string rootPath)
        {
            var objPowerPoint = new Application();
            foreach (var file in Directory.GetFiles(rootPath, "*.ppt?", SearchOption.AllDirectories))
            {
                if (file.Contains("reference")) continue;
                var inputPath = Path.GetFullPath(file);
                var dir_name = Path.GetDirectoryName(inputPath);
                var file_name = Path.GetFileNameWithoutExtension(inputPath);
                var outputPath = Path.Combine(dir_name, file_name + ".index.txt");
                // skipping documents that have been already converted
                if (File.Exists(outputPath) && File.GetLastWriteTime(outputPath) >= File.GetLastWriteTime(inputPath))
                {
                    Console.WriteLine($"skipping {inputPath}");
                    continue;
                }
                else
                {
                    Console.WriteLine($"Extracting index for {inputPath}");
                    var pptDoc = objPowerPoint.Presentations.Open(inputPath);
                    using (var writer = new StreamWriter(outputPath))
                    {
                        for (int i = 1; i <= pptDoc.Slides.Count; ++i)
                        {
                            var slide = pptDoc.Slides[i];
                            var title = slide.Shapes.Title.TextFrame.TextRange.Text
                                        .Replace("\v", "")
                                        .Replace("\t", "")
                                        .Replace("\r\n", "");
                            var pageNum = slide.SlideNumber;
                            writer.WriteLine($"{title}\t{pageNum}");
                        }
                    }
                    pptDoc.Close();
                    Console.WriteLine($"Completed: {outputPath}.");
                }
            }
        }
        static void Main(string[] args)
        {
            GenerateIndex(args[0]);
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