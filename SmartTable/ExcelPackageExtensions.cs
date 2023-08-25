using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartWriter
{
    public static class ExcelPackageExtensions
    {
        public static string TMP_DIRECTORY = @"C:\tmp\";

        public static ExcelPackage Merge(this IEnumerable<ExcelPackage> packages)
        {
            var masterPackage = new ExcelPackage(new FileInfo(TMP_DIRECTORY + @"firstExcel.xlsx"));

            int i = 1;

            foreach (var package in packages)
            {
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    var omonymousSheet = masterPackage.Workbook.Worksheets.SingleOrDefault(s => sheet.Name == s.Name);

                    string worksheetName = omonymousSheet == null ? sheet.Name :
                                           string.Format("{0}_{1}", sheet.Name, DateTime.Now.ToString("yyyyMMddhhssmmm") + i++);

                    masterPackage.Workbook.Worksheets.Add(worksheetName, sheet);
                }
            }

            return masterPackage;
        }

        public static void SaveExcel(this ExcelPackage package, string filepath, bool autoFitColumns = true)
        {
            EnsureFilePathExists(filepath);

            if (autoFitColumns)
            {
                foreach (var ws in package.Workbook.Worksheets)
                {
                    ws.Cells.AutoFitColumns();
                }
            }

            using (FileStream fs = new FileStream(filepath, FileMode.Create))
            {
                package.SaveAs(fs);
            }
        }

        public static void SavePdf(this ExcelPackage package, string filepath, string excelFilepath = null, string tmpDirectory = null)
        {
            var tmpSaveNeeded = string.IsNullOrWhiteSpace(excelFilepath);

            tmpDirectory = tmpDirectory ?? TMP_DIRECTORY;

            if (tmpSaveNeeded)
            {
                Directory.CreateDirectory(tmpDirectory);

                excelFilepath = tmpDirectory + Path.GetFileNameWithoutExtension(filepath) + ".xlsx";

                if (File.Exists(excelFilepath))
                {
                    File.Delete(excelFilepath);
                }

                SaveExcel(package, excelFilepath);
            }

            SavePdf(filepath, excelFilepath);

            if (tmpSaveNeeded)
            {
                File.Delete(excelFilepath);
            }
        }

        private static void SavePdf(string filepath, string tmpFilepath)
        {
            var MSDoc = new Application()
            {
                Visible = false,
                DisplayAlerts = false
            };

            MSDoc.Workbooks
                .Open(tmpFilepath)
                .ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filepath);

            MSDoc.Workbooks.Close();
        }

        private static void EnsureFilePathExists(string filepath)
        {
            var dir = Path.GetDirectoryName(filepath);

            Directory.CreateDirectory(dir);

            if (!File.Exists(filepath))
            {
                File.Create(filepath)
                    .Close();
            }
        }
    }
}