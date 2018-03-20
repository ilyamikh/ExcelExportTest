using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExportTest
{
    class LoadExcel
    {
        public static List<MenuItem> item_catalogue = new List<MenuItem>();
        public static List<Ingredient> ingredient_catalogue = new List<Ingredient>();

        private static Excel._Application excelItemList = new Excel.Application();
        private static Excel._Application excelIngredientList = new Excel.Application();
        public static Excel._Application excelEditCheck = new Excel.Application();

        private static Excel._Worksheet itemSheet;
        private static Excel._Worksheet ingredientSheet;
        public static Excel._Worksheet EditCheck;

        public static void loadSheets()
            {
                excelItemList.Workbooks.Open("C:/Users/Ella/My Documents/Production Generator Raw Data/item_list_raw.xlsx");
                excelIngredientList.Workbooks.Open("C:/Users/Ella/My Documents/Production Generator Raw Data/ingredient_list_raw.xlsx");
                excelEditCheck.Workbooks.Open("C:/Users/Ella/My Documents/claims/SNP 1516/NOV 15/Snack November 2015.xls");

                itemSheet = excelItemList.ActiveSheet;
                ingredientSheet = excelIngredientList.ActiveSheet;
                EditCheck = excelEditCheck.ActiveSheet; //what will it open if the active sheet is not the one I need?
            }
            
        public static void createItemList ()
        {
            Excel.Range itemIDRange;
            Excel.Range itemNameRange;
            Excel.Range ingredientIDRange;

            //magic number, fix later
            for (int i = 1; i < 60; i++)
            {
                string idCell = "A" + (i + 1);
                string nameCell = "B" + (i + 1);
                string ingredientCell = "C" + (i + 1);

                itemIDRange = itemSheet.get_Range(idCell);
                itemNameRange = itemSheet.get_Range(nameCell);
                ingredientIDRange = itemSheet.get_Range(ingredientCell);

                addItemToCatalog(Convert.ToInt32(itemIDRange.Value),
                                    itemNameRange.Text,
                                    ingredientIDRange.Text);
            }

        }

        public static void createIngredientList()
        {
            Excel.Range ingredientIDRange;
            Excel.Range ingredientVarNameRange;
            Excel.Range ingredientNameRange;
            Excel.Range servingSizeRange;
            Excel.Range descriptionRange;
            Excel.Range packageRange;

            //another magic number, fix later
            for (int i = 1; i < 112; i++)
            {
                string IDcell = "A" + (i + 1);
                string varNameCell = "B" + (i + 1);
                string nameCell = "C" + (i + 1);
                string servSizeCell = "D" + (i + 1);
                string descriptionCell = "E" + (i + 1);
                string packageCell = "F" + (i + 1);

                ingredientIDRange = ingredientSheet.get_Range(IDcell);
                ingredientVarNameRange = ingredientSheet.get_Range(varNameCell);
                ingredientNameRange = ingredientSheet.get_Range(nameCell);
                servingSizeRange = ingredientSheet.get_Range(servSizeCell);
                descriptionRange = ingredientSheet.get_Range(descriptionCell);
                packageRange = ingredientSheet.get_Range(packageCell);

                addIngredientToCatalog(
                            Convert.ToInt32(ingredientIDRange.Value),
                            ingredientVarNameRange.Text,
                            ingredientNameRange.Text,
                            Convert.ToDouble(servingSizeRange.Value),
                            descriptionRange.Text,
                            Convert.ToDouble(packageRange.Value)
                    );


            }
        }

        private static void addItemToCatalog(int inID, string inName, string inList)
        {
            MenuItem temp = new MenuItem(inID, inName, inList);
            item_catalogue.Add(temp);
        }

        private static void addIngredientToCatalog(int inID, string inVar, string inName, double inSize, string inDescription, double inPackage)
        {
            Ingredient temp = new Ingredient(inID, inVar, inName, inSize, inDescription, inPackage);
            ingredient_catalogue.Add(temp);
        }

        public static void closeApp()
        {
            excelIngredientList.Quit();
            if (excelIngredientList != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelIngredientList);
            if (ingredientSheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ingredientSheet);

            excelIngredientList = null;
            ingredientSheet = null;

            excelItemList.Quit();
            if (excelItemList != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelItemList);
            if (itemSheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(itemSheet);

            excelItemList = null;
            itemSheet = null;

            excelEditCheck.Quit();
            if (excelEditCheck != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelEditCheck);
            if (EditCheck != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(EditCheck);

            excelEditCheck = null;
            EditCheck = null;

            GC.Collect();
        }
    }
}
