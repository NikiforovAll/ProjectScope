using System;
using XMindAPI.LIB;

namespace ProjectScope.Client
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileExtension = "xmind";
            string defaultSheetName = "projectScope";
            string dataSource = @"C:\Users\HYS\Downloads\internal_actionlist.xlsx";
            XMindWorkBook book = new XMindWorkBook($"test.{fileExtension}");
            ConfigureXMindWorkBook(book);
            book.AddSheet(defaultSheetName);
            book.Save();
        }
        public static void ConfigureXMindWorkBook(XMindWorkBook book)
        {

        }
    }
}
    