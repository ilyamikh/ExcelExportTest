using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExportTest
{
    class Program
    {
       

        static void Main(string[] args)
        {
            LoadExcel.loadSheets();
            LoadExcel.createItemList();
            LoadExcel.createIngredientList();
            //int choice = 0;
            //do{
            Meal testMeal = Interface.buildMeal();
            //Interface.showMealData(testMeal);
            Export.buildSheet(testMeal);
            
            //    Console.WriteLine("Use meal again? 1 - Yes, 2 - No");
            //    choice = Interface.getChoice(1, 2);
            //    if (choice == 1)
            //    {

            //        LoadExcel.closeApp();
                   
            //        Console.WriteLine("Enter date:");
            //        testMeal.setDate(Console.ReadLine());
            //        Export.buildSheet(testMeal);
            //    }
            
            //}while(choice != 2);

            LoadExcel.closeApp();

        //    List<Meal> meals = new List<Meal>()
        //    {
        //    new Meal("Borsch", "Pilaf", "Vegetable Salad"),
        //    new Meal("Vermicelli Bullion", "Pelmeni", "Tomato and Corn"),
        //    new Meal("Buckwheat Soup", "Macaroni and Chicken", "Pickles")
        //    };

        //    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        //    excel.Workbooks.Add();
        //    Microsoft.Office.Interop.Excel._Worksheet workSheet = excel.ActiveSheet;

        //    try
        //    {
        //        workSheet.Cells[1, "A"] = "Soup";
        //        workSheet.Cells[1, "B"] = "Main Course";
        //        workSheet.Cells[1, "C"] = "Salad";

        //        int row = 2;
        //        foreach (Meal current in meals)
        //        {
        //            workSheet.Cells[row, "A"] = current.getSoup();
        //            workSheet.Cells[row, "B"] = current.getMainCourse();
        //            workSheet.Cells[row, "C"] = current.getSalad();

        //            row++;
        //        }

        //        string filename = "testSheet2.xlsx";
        //        workSheet.SaveAs(filename);
        //        Console.WriteLine("Success!");

        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("Error");
        //        Console.WriteLine(e.Message);
        //    }
        //    finally
        //    {
        //        excel.Quit();
        //        if (excel != null)
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        //        if (workSheet != null)
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet);

        //        excel = null;
        //        workSheet = null;

        //        GC.Collect();
        //    }

        }
    }
}
