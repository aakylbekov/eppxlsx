using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace ConsoleApplication1
{
    class Program
    {
        static Model1 db = new Model1();

        static void Main(string[] args)
        {
            ExcelPackage exp = new ExcelPackage();
            ExcelWorksheet worksheet = exp.Workbook.Worksheets.Add("List1");

            int row = 2;
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Column(2).Width = 50;
            worksheet.Cells[1, 3].Value = "IP";

            foreach (Area area in db.Area)
            {
                worksheet.Cells[row, 1].Value = area.AreaId;
                worksheet.Cells[row, 2].Value = area.FullName;
                worksheet.Cells[row, 3].Value = area.IP;
                row++;
            }

            Dictionary<string, Area> dicIP = db.Area.
                Where(w => !string.IsNullOrEmpty(w.IP) && w.ParentId != 0).
                Select(s =>new {s.IP}).
                Distinct().
                Select(s=>new { ip= s.IP, area = db.Area.FirstOrDefault(f=>f.IP==s.IP)}).
                ToDictionary(d => d.ip, d => d.area);

            ExcelWorksheet worksheet2 = exp.Workbook.Worksheets.Add("List2");
            row = 2;
            foreach (var item in dicIP)
            {
                worksheet2.Cells[row, 1].Value = item.Key;
                worksheet2.Cells[row, 2].Value = item.Value.FullName;            
            }
            ILookup<string, Area> lkp = db.Area.ToLookup(l => l.IP, l => l);
            FileStream fs = File.Create("Excelca.xlsx");
            exp.SaveAs(fs);


        }
    }
}
