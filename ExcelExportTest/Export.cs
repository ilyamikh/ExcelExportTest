using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExportTest
{
    class Export
    {
        private static List<Ingredient> production_list = new List<Ingredient>();
        private static Excel.Application export = new Excel.Application();
        private static Excel._Worksheet currentSheet;

        private static void prepareSheet()
        {       
            export.Workbooks.Add();
            currentSheet = export.ActiveSheet;
        }

        /*
         * This function returns the row number in the EditCheck excel with the date of the meal
         * to be used later to get the actual counts from the EditCheck
         * The dates are on column A, starting on row 7 and end on row 37
         */ 
        private static int getRowNumber(String date)
        {
            int dateRow = -1; //let's have everything crash if it can't find the date, why not
            Excel.Range currentDateRange;
            String cell;

            int rowStart = 7;
            int rowEnd = 37;
            DateTime mealDate = Convert.ToDateTime(date);

            for (int currentRow = rowStart; currentRow <= rowEnd; currentRow++ ) {
                cell = "A" + currentRow;
                currentDateRange = LoadExcel.EditCheck.get_Range(cell);
                if (currentDateRange.Value != null) //eliminate all the blank cells
                {

                    //DateTime excelDate = Convert.ToDateTime(currentDateRange.Text);
                    //Console.WriteLine("The value " + currentDateRange.Text + " in the excel sheet is " + excelDate);

                    if (Convert.ToDateTime(currentDateRange.Text) == mealDate) dateRow = currentRow; //here's hoping string to datetime conversion works! (it does)
                }
            }
            
                return dateRow;

        }

        private static int getCellValueInt(String column, int row) {
            string idCell = column + row;
            Excel.Range cell;
            cell = LoadExcel.EditCheck.get_Range(idCell);
            return Convert.ToInt32(cell.Value);
        }

        //add logic to round up to nearest 5 or 10
        private static int roundUp(int num)
        {
            double deci = num / 10;
            int val = (int)deci;
            return (val * 10) + 10;

        }

        public static void buildSheet(Meal currentMeal)
        {
            prepareSheet();
            int rowNumber = getRowNumber(currentMeal.getDate());

            //Console.WriteLine("Enter Student Servings.");
            //int studentServings = Interface.enterNumber();
            //Console.WriteLine("Enter Adult Servings.");
            //int adultServings = Interface.enterNumber();
            //Console.WriteLine("Enter Reimburseable Meals.");
            //int reimburseable = Interface.enterNumber();
            //Console.WriteLine("Enter Non-Reimburseable Meals.");
            //int nonReimburseable = Interface.enterNumber();
            string filename = "C:\\Users\\Ella\\Desktop\\Production1516\\SNACK\\NOV 2015\\" + currentMeal.getDate() + currentMeal.getName();
            Console.WriteLine("Output will be saved to " + filename);
           
            /*      B   Paid Claimed
             *      E   Free Claimed
             *      H   Reduced Price Claimed
             *      
             */
            int adultServings = 20; //yes
            int paidClaimed = getCellValueInt("B", rowNumber);
            int freeClaimed = getCellValueInt("E", rowNumber);
            int reducedClaimed = getCellValueInt("H", rowNumber);

            int studentServings = roundUp(paidClaimed + freeClaimed + reducedClaimed);
            int totalServings = studentServings + adultServings;
            int reimburseable = freeClaimed + reducedClaimed;
            int nonReimburseable = paidClaimed + adultServings; //super advanced food program math

            currentSheet.Cells[1, "A"] = "Menu Production Record";
            currentSheet.Cells[1, "H"] = "All values in grams and milliliters";
            currentSheet.Cells[2, "A"] = "School";
            currentSheet.Cells[2, "B"] = "RLES";
            currentSheet.Cells[3, "A"] = currentMeal.getName();
            currentSheet.Cells[4, "A"] = currentMeal.getDate();

            currentSheet.Cells[6, "A"] = "Meals Served";
            currentSheet.Cells[7, "A"] = "Children";
            currentSheet.Cells[8, "A"] = "Adults";
            currentSheet.Cells[9, "A"] = "Total";

            currentSheet.Cells[7, "B"] = studentServings;
            currentSheet.Cells[8, "B"] = adultServings;
            currentSheet.Cells[9, "B"] = totalServings;

            currentSheet.Cells[2, "D"] = "Menu:";

            for (int i = 0; i < currentMeal.getCount(); i++)
            {
                int row = i + 2;
                currentSheet.Cells[row, "E"] = currentMeal.getItem(i).getName();
            }

            int currentItemRow = 11; //default magic number value, see excel
            //if there are more items that can fit with the default frame, adjust item row
            if (currentMeal.getCount() > 8)
                currentItemRow = currentMeal.getCount() + 2;

            postIngredients(currentItemRow, currentMeal, studentServings, adultServings, reimburseable, nonReimburseable);

            currentSheet.SaveAs(filename);
            cleanUp();
        }

        private static void postIngredients(int currentRow, Meal currentMeal, int studentServings, int adultServings, int reimburseable, int nonReimburseable)
        {
            int totalServings = studentServings + adultServings;
            int currentItemRow = currentRow;

            for (int i = 0; i < currentMeal.getCount(); i++) // loop where all the work is done, itirates through all menu items
            {
                //item header
                currentSheet.Cells[currentItemRow, "A"] = "Name";
                currentSheet.Cells[currentItemRow, "B"] = currentMeal.getItem(i).getName();
                currentSheet.Cells[currentItemRow, "C"] = "ID";
                currentSheet.Cells[currentItemRow, "D"] = currentMeal.getItem(i).getID();
                currentSheet.Cells[currentItemRow, "E"] = "Ingredient Count";
                currentSheet.Cells[currentItemRow, "G"] = currentMeal.getItem(i).getIngredientCount();
                //ingredient column headers
                currentItemRow++;
                
                currentSheet.Cells[currentItemRow, "A"] = "Ingredient Name";
                currentSheet.Cells[currentItemRow, "C"] = "ID";
                currentSheet.Cells[currentItemRow, "D"] = "Package";
                currentSheet.Cells[currentItemRow, "H"] = "Students";
                currentSheet.Cells[currentItemRow, "I"] = "Total";
                currentSheet.Cells[currentItemRow, "G"] = "Portion";
                currentSheet.Cells[currentItemRow, "J"] = "Prepared";
                currentSheet.Cells[currentItemRow, "K"] = "Reimburseable";
                currentSheet.Cells[currentItemRow, "L"] = "Non-Reimburseable";
                currentSheet.Cells[currentItemRow, "M"] = "Leftovers";

                currentItemRow++;
                MenuItem currentItem;
                currentItem = currentMeal.getItem(i);

                for (int j = 0; j < currentItem.getIngredientCount(); j++) //loop that posts ingredient info, uses currentIngredientRow
                {
                    currentSheet.Cells[currentItemRow, "A"] = LoadExcel.ingredient_catalogue[currentItem.getIngredientId(j) - 1].getName();
                    currentSheet.Cells[currentItemRow, "C"] = LoadExcel.ingredient_catalogue[currentItem.getIngredientId(j) - 1].getID();
                    currentSheet.Cells[currentItemRow, "D"] = LoadExcel.ingredient_catalogue[currentItem.getIngredientId(j) - 1].getDescription();
                    currentSheet.Cells[currentItemRow, "H"] = studentServings;
                    currentSheet.Cells[currentItemRow, "I"] = totalServings;
                    currentSheet.Cells[currentItemRow, "G"] = LoadExcel.ingredient_catalogue[currentItem.getIngredientId(j) - 1].getServingSize();
                    currentSheet.Cells[currentItemRow, "J"] = LoadExcel.ingredient_catalogue[currentItem.getIngredientId(j) - 1].getServingSize() * totalServings;
                    currentSheet.Cells[currentItemRow, "K"] = reimburseable;
                    currentSheet.Cells[currentItemRow, "L"] = nonReimburseable;
                    currentSheet.Cells[currentItemRow, "M"] = totalServings - (reimburseable + nonReimburseable);
                    
                    currentItemRow++;
                }

            }
        }

        private static void cleanUp()
        {
            export.Quit();
            if (export != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(export);
            if (currentSheet != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(currentSheet);

            export = null;
            currentSheet = null;

            GC.Collect();
        }

        private static void addIngredient(MenuItem currentItem)
        {

        }

    }
}
