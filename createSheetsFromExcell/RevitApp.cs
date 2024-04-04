using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace createSheetsFromExcell
{
    internal class RevitApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string tabeName = "EExcell";
            string panelName = "Sheets";
            string ButtonName = "SheetsCreator";
            string ButtonInternalName = "SheetsCreator_Btn";
            string assemplyName = "createSheetsFromExcell";
            string className = "CreateSheets";
            string imageName = "Treetog-I-Documents.ico";
            string path = Assembly.GetExecutingAssembly().Location;

            #region CreateUIButton
            //01-Create tab
            application.CreateRibbonTab(tabeName);
            //02-Create Panael
            var myPanel = application.CreateRibbonPanel(tabeName, panelName);
            //03-Create Push Button
            PushButtonData myButtonData = new PushButtonData(ButtonInternalName, ButtonName, path, $"{assemplyName}.{className}");
            PushButton myButton = myPanel.AddItem(myButtonData) as PushButton;

            Uri myUri = new Uri($"pack://application:,,,/{assemplyName};component/Image/{imageName}");
            BitmapImage myImage = new BitmapImage(myUri);
            myButton.LargeImage = myImage; 
            #endregion
            return Result.Succeeded;
        }
    }
}
