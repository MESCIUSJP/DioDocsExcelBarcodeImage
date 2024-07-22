// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;

Console.WriteLine("ワークシートのセルにバーコード・QRコードを追加する");

// 新規ワークブックの作成
var workbook = new Workbook();
IWorksheet worksheet = workbook.Worksheets[0];

// セルのサイズを調整
worksheet.Range["A:B"].ColumnWidth = 20;
worksheet.Range["B1"].RowHeight = 60;
worksheet.Range["B2"].RowHeight = 60;

// 表示位置を調整
worksheet.Range["A1:B2"].HorizontalAlignment = HorizontalAlignment.Center;
worksheet.Range["A1:B2"].VerticalAlignment = VerticalAlignment.Center;

// バーコードで利用する値
worksheet.Range["A1"].Value = "692031229621";
worksheet.Range["A2"].Value = "メシウス株式会社";

// バーコードを設定
worksheet.Range["B1"].Formula = "=BC_EAN13(A1)";
worksheet.Range["B2"].Formula = "=BC_QRCODE(A2)";

// バーコードを画像に変換
workbook.ConvertBarcodeToPicture(ImageType.PNG);

// ファイルに出力
workbook.Save("result.xlsx");
workbook.Save("result.pdf");

