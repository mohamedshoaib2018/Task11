using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace createSheetsFromExcell
{
    internal class SheetData
    {
        //Attributes
        string _sheetName = null;
        string _sheetNum = null;


        //Properties
        public string SheetName { get => _sheetName; set => _sheetName = value; }
        public string SheetNum { get => _sheetNum; set => _sheetNum = value; }


        //Constructor
        public SheetData(string sheetName, string sheetNum)
        {
            SheetName = sheetName;
            SheetNum = sheetNum;
        }

    }
}
