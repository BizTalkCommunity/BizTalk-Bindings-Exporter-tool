using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace BizTalk_Exporter
{
    public class excelData
    {
        public string portName;
        public string portURI;
        public excelData(string _portName, string _portURI)
        {
            portName = _portName;
            portURI = _portURI;
        }
    }
    public class ExcelOperations
    {
        /// <summary>
        /// Reads all the cells in the given Excel File.
        /// </summary>
        /// <param name="environment">QA or PROD</param>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public List<excelData> ReadExcelFile(string environment, string excelFile)
        {
            try
            {
                var package = new ExcelPackage(new System.IO.FileInfo(excelFile));
                ExcelWorksheet sheet = package.Workbook.Worksheets[environment];
                if (sheet == null)
                    throw new Exception("Could not find Sheet");
                int startRow = sheet.Dimension.Start.Row;
                int endRow = sheet.Dimension.End.Row;
                List<excelData> portsList = new List<excelData>();
                for (int i = startRow; i <= endRow; i++)
                {
                    portsList.Add(new excelData(
                        sheet.Cells[i, 1].Text,
                        sheet.Cells[i, 5].Text
                    ));
                }
                return portsList;
            }
            catch (Exception ex)
            { throw ex; }
        }
    }
}
