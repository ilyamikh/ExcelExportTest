using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportTest
{
    class MenuItem
    {
        private int ITEM_ID;
        private string itemName;
        private int[] INGREDIENT_ID;

        public MenuItem(int ID, string name, string ingredientList)
        {
            ITEM_ID = ID;
            itemName = name;
            INGREDIENT_ID = new int[listToArray(ingredientList).Length];
            //assign the ingredient id's
            for (int i = 0; i < INGREDIENT_ID.Length; i++)
                INGREDIENT_ID[i] = listToArray(ingredientList)[i];
        }

        public int getID() {
            return ITEM_ID;
        }

        public string getName()
        {
            return itemName;
        }

        public int getIngredientId(int index)
        {
            return INGREDIENT_ID[index];
        }

        public int getIngredientCount()
        {
            return INGREDIENT_ID.Length;
        }
        /*
         * Converts the input string (list of numbers representing ingredient codes) into integer array
         */
        public int[] listToArray(string numberList) {
            string[] stringArray = numberList.Split(' ');
            int[] intArray = new int[stringArray.Length];
            for (int i = 0; i < stringArray.Length; i++)
            {
                intArray[i] = Convert.ToInt32(stringArray[i]);
            }
            return intArray;
        }
    }
}
