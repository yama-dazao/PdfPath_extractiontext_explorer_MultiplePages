using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using UglyToad.PdfPig;

class PdfPathExtractionTextToFile
{
    [STAThread]
    static void Main(string[] args)
    {
        // ファイルダイアログの設定
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Title = "PDFファイルを選択してください",
            Filter = "PDFファイル (*.pdf)|*.pdf",
            InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            string pdfPath = openFileDialog.FileName;
            string downloadsPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
            string outputFilePath = Path.Combine(downloadsPath, "ExtractedText.txt");

            try
            {
                using (var pdf = PdfDocument.Open(pdfPath))
                {
                    using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                    {
                        writer.WriteLine("--- PDFから抽出した内容 ---\n");

                        int totalPages = 0; // ページ数カウント用
                        foreach (var page in pdf.GetPages())
                        {
                            try
                            {
                                writer.WriteLine($"ページ {page.Number}:\n");
                                writer.WriteLine(page.Text);
                                writer.WriteLine("\n-------------------\n");
                                totalPages++;
                            }
                            catch (Exception pageEx)
                            {
                                Console.WriteLine($"ページ {page.Number} の処理中にエラーが発生しました: {pageEx.Message}");
                            }
                        }

                        writer.WriteLine($"総ページ数: {totalPages} ページ\n");
                    }
                }

                Console.WriteLine("PDFの内容をテキストファイルに書き出しました。");
                Console.WriteLine($"保存先: {outputFilePath}");
                System.Diagnostics.Process.Start("notepad.exe", outputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("キャンセルされました。プログラムを終了します。");
        }
    }
}
