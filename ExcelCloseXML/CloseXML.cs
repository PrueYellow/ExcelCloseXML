using ClosedXML.Excel;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelCloseXML
{
    internal class CloseXML
    {
        static void Main(string[] args)
        {
            //建立活頁簿
            IXLWorkbook wb = new XLWorkbook(@"C:\Users\G-pro\Downloads\BookProductsAfter.xlsx");


            //建立工作表
            var worksheet = wb.Worksheet(1);

            //讀取元素

            

            int columnCount = 7; // 列數:書名 出版社...有7個欄位

            int rowCount = worksheet.RowsUsed().Count(); // 根據Excel中取得實際行數

            for (int row = 2; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    var basicdata = worksheet.Cell(1, column).Value.ToString();

                    //抓取時間位置並轉格式
                    string cellValue = column == 3 ? worksheet.Cell(row, column).GetDateTime().ToString("yyyy/MM/dd") : worksheet.Cell(row, column).GetString();

                    var fuckingGay = $"{basicdata} {cellValue}";

                    Console.WriteLine(fuckingGay);
                }
                Console.WriteLine();
            }




        }

    }
}