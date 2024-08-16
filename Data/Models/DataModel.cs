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

        public DataModel()
        {
            InitColumns = new List<string>();
            CalculatedColumns = new List<string>();
        }
    }

}
