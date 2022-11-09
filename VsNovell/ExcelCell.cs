using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelForms3.Consult
{
    internal class ExcelCell
    {
        public int Cell { get; set; }
        public string StartElement { get; private set; }
        public string EndElement { get; set; }
        public bool Extract { get; }

        public ExcelCell(int cell)
        {
            Cell = cell;
        }

        public ExcelCell(int cell, string startElement, string endElement)
        {
            Cell = cell;
            StartElement = startElement;    
            EndElement = endElement;
            Extract = true;
        }
    }
}
