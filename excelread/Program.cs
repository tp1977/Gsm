using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excelread;

namespace excelread
{
    class Program
    {
       
        static void Main(string[] args)
        {
            ExcelReaderFile ex = new ExcelReaderFile("D:\\excel\\test.xlsx");

            string cell = ex.getCellData("test", 1, 2);

            int rcount = ex.getRowCount("test");
            int ccount = ex.getColumnCount("test");

            Console.WriteLine(rcount);
            Console.WriteLine(ccount);


            Console.WriteLine(cell);

            Console.WriteLine("I am adding this line");
            Console.WriteLine("I am adding second line");
        }
    }
}
