using System;

namespace OpenXMLSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report();
            
            report.CreateExcelDoc(@"C:\Users\12499\Desktop\info.xlsx");
            report.InsertText(@"C:\Users\12499\Desktop\info.xlsx", "My Name is Bond, James Bond.");
        }
    }
}
