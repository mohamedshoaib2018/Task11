using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace createSheetsFromExcell
{
    [Transaction(TransactionMode.Manual)]
    public class CreateSheets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;


            //00-Variables
            string filePath = null;
            List<SheetData> dataList = new List<SheetData>();
            ViewSheet mySheet = null;

            //01-Open File
            OpenFileDialog myDialog = new OpenFileDialog()
            {
                Filter = "Excell Files|*.xls;*.xlsx;*.xlsb;*.csv;*.xlsm", Multiselect = false
                
            };  
            bool isOpend = (bool)myDialog.ShowDialog();
            
            //02-Get file Path
            if (isOpend == true)
            {
                filePath = myDialog.FileName; 
            }

            ExcellHandeller handeller = new ExcellHandeller();
            dataList = handeller.get_Data_From_Excell(filePath);


            #region Create sheet
            var titelBlockType = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsElementType()
                .FirstElementId();
            using (Transaction trans = new Transaction(doc,"Create sheets"))
            {
                trans.Start();
                foreach (SheetData item in dataList)
                {
                    mySheet = ViewSheet.Create(doc, titelBlockType);
                    mySheet.SheetNumber = item.SheetNum;
                    mySheet.Name = item.SheetName;
                }
                trans.Commit(); 
            }
            

            #endregion

            return Result.Succeeded;
        }
    }
}
