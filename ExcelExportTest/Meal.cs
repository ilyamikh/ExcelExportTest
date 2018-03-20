using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExportTest
{
    class Meal
    {
        private string date;
        private string name;
        private List<MenuItem> items = new List<MenuItem>();

        public Meal(string inName, string inDate)
        {
            name = inName;
            date = inDate;
        }

        public void addItem(MenuItem newItem)
        {
            items.Add(newItem);
        }

        public int getCount()
        {
            return items.Count;
        }

        public string getName()
        {
            return name;
        }

        public string getDate()
        {
            return date;
        }
        public void setDate(string newDate)
        {
            date = newDate;
        }

        public MenuItem getItem(int id)
        {
            return items[id];
        }
    }
}
