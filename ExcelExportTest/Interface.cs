using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportTest
{
    class Interface
    {
        public static int getChoice(int start, int totalChoices)
        {
            int choice = start - 1;
            try
            {
                Console.WriteLine("Enter Choice: ");
                string input = Console.ReadLine();
                if (!Int32.TryParse(input, out choice) || !(choice >= start && choice <= totalChoices))
                    throw new Exception("Invalid Input");

            }
            catch (Exception invalid)
            {
                Console.WriteLine(invalid.Message);
                getChoice(start, totalChoices);
            }
            return choice;
        }

        private static int getItemChoice()
        {
            showItems(LoadExcel.item_catalogue.Count);
            return getChoice(1, LoadExcel.item_catalogue.Count);

        }

        private static string enterMealType()
        {
            Console.Write("Enter Meal Type: ");
            return Console.ReadLine();
        }

        private static string enterDate()
        {
            Console.Write("Enter Date: ");
            return Console.ReadLine();
        }

        public static int enterNumber()
        {
            int students = -1;
            try
            {
                Console.Write("Enter number: ");
                string input = Console.ReadLine();
                if (!Int32.TryParse(input, out students) || students < 0)
                    throw new Exception("Invalid Input");
            }
            catch (Exception invalid)
            {
                Console.Write(invalid.Message);
                enterNumber();
            }
            return students;
        }
        
        private static void showItems(int totalItems)
        {
            for (int id = 0; id < totalItems; id++)
            {
                    Console.WriteLine((id + 1) + ": " + LoadExcel.item_catalogue[id].getName());
            }
        }


        public static Meal buildMeal()
        {
            //string type = enterMealType();

            string type = "SNACK";
            Console.WriteLine(type);
            string date = enterDate();
            
            Meal currentMeal = new Meal(type, date);

       
            int choice = 0;
            do
            {
                Console.WriteLine("1: Add Item");
                Console.WriteLine("2: Done");
                choice = getChoice(1, 2);
                if (choice == 1)
                {
                    currentMeal.addItem(LoadExcel.item_catalogue[getItemChoice() - 1]);
                }
            } while (choice != 2);

            return currentMeal;
        }

        public static void showMealData(Meal testMeal)
        {
            Console.WriteLine("Meal type: " + testMeal.getName());
            Console.WriteLine("Date: " + testMeal.getDate());
            Console.WriteLine("Items in Meal: " + testMeal.getCount());
            for (int i = 0; i < testMeal.getCount(); i++)
            {
                Console.WriteLine("Item " + i + ": " + testMeal.getItem(i).getName());
                Console.WriteLine("Ingredient Count: " + testMeal.getItem(i).getIngredientCount());
                showIngredientData(testMeal.getItem(i));               
            }
        }

        private static void showIngredientData(MenuItem testItem)
        {
            for (int i = 0; i < testItem.getIngredientCount(); i++)
            {
                Console.WriteLine("ID: " + LoadExcel.ingredient_catalogue[testItem.getIngredientId(i) - 1].getID());
                Console.WriteLine("Name: " + LoadExcel.ingredient_catalogue[testItem.getIngredientId(i) - 1].getName());
                Console.WriteLine("Serving Size: " + LoadExcel.ingredient_catalogue[testItem.getIngredientId(i) - 1].getServingSize());
                Console.WriteLine("Package Description: " + LoadExcel.ingredient_catalogue[testItem.getIngredientId(i) - 1].getDescription());
                Console.WriteLine("Package Size: " + LoadExcel.ingredient_catalogue[testItem.getIngredientId(i) - 1].getPackage());
            }
        }
    }
}