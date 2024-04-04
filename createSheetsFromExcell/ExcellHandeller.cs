using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Animation;

namespace createSheetsFromExcell
{
    internal class ExcellHandeller
    {
        //Attributes
        List<SheetData>_sheetDataList = new List<SheetData>();
        //Contructor
        public ExcellHandeller(){ }

        //Methods
        /// <summary>
        /// This Methods Get Data From Excell File
        /// </summary>
        /// <param name="filepath">The Path Of The Excell File</param>
        public List<SheetData> get_Data_From_Excell(string filepath)
        {
            _sheetDataList.Clear();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage myPackage = new ExcelPackage(new System.IO.FileInfo(filepath)))
            {
                var mySheet = myPackage.Workbook.Worksheets[0];

                
                for (int i = 2; i < mySheet.Rows.Count(); i++)
                {
                    string strSheetNum = mySheet.Cells[i, 1].Value.ToString();
                    string strSheetName = mySheet.Cells[i, 2].Value.ToString();
                    SheetData mySheetData = new SheetData(strSheetName, strSheetNum);
                    _sheetDataList.Add(mySheetData);
                }
                return _sheetDataList;

            }
        }
    }
}
