using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelForms3.Consult
{
    public class ExcelData
    {
        public List<string[]> Data { get; set; }

        public ExcelData()
        {
            Data = new List<string[]>(); 
        }
    }
}
