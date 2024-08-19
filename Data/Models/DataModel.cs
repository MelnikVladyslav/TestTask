using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data.Models
{
    public class DataModel
    {
        public string ID { get; set; }

        public List<string> InitColumns { get; set; }

        public List<string> CalculatedColumns { get; set; }
        public List<string> UserColumns { get; set; }
        public List<string> Rows { get; set; }

        public DataModel()
        {
            InitColumns = new List<string>();
            CalculatedColumns = new List<string>();
            UserColumns = new List<string>();
            Rows = new List<string>();
        }

        public DataModel(int numberOfColumns)
        {
            InitColumns = new List<string>(new string[numberOfColumns]);
            CalculatedColumns = new List<string>(new string[numberOfColumns]);
            UserColumns = new List<string>(new string[numberOfColumns]);
            Rows = new List<string>(new string[numberOfColumns]);
        }
    }

}
