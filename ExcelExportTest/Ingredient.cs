using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportTest
{
    class Ingredient
    {
        private int INGREDIENT_ID;
        private string VAR_NAME; //legacy
        private string name;
        private double servingSize;
        private string packageDescription;
        private double packageSize;

        public Ingredient(int inID, string inVarName, string inName, double inServSize, string inDescription, double inPackage)
        {
            INGREDIENT_ID = inID;
            VAR_NAME = inVarName;
            name = inName;
            servingSize = inServSize;
            packageDescription = inDescription;
            packageSize = inPackage;

        }

        public int getID()
        {
            return INGREDIENT_ID;
        }
        public string getVarName()
        {
            return VAR_NAME;
        }
        public string getName()
        {
            return name;
        }
        public double getServingSize()
        {
            return servingSize;
        }
        public string getDescription()
        {
            return packageDescription;
        }
        public double getPackage()
        {
            return packageSize;
        }
    }
}
