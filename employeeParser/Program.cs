
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace employeeParser
{
    class Program
    {
        static void Main(string[] args)
        {

            List<int> nums = new List<int>() { 2, 3, 4, 6, 9 };

            Parser p = new Parser(@"C:\Users\14168\Desktop\EmployeeMatrix.xlsx");
            p.CreateExcelPackage("Four",nums,2,20);


            var ans=p.GetDict();
            p.GetFormattedList();
            p.ExportToTxtFile(@"C:\Users\14168\Desktop\EmployeeMatrix.txt");
           
            Console.ReadKey();




        }


        //        string filePath = @"C:\Users\14168\Desktop\EmployeeMatrix.xlsx";



        //        FileInfo fi = new FileInfo(filePath);
        //            using (ExcelPackage excelPackage = new ExcelPackage(fi))
        //            {
        //                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets["Three"];

        //    List<object> columnB = new List<object>();

        //                for (int i = 2; i<56; i++)
        //                {
        //                   columnB.Add(firstWorksheet.Cells[i, 2].Value);
        //                }


        //List<object> columnC = new List<object>();

        //                for (int i = 2; i< 56; i++)
        //                {
        //                    columnC.Add(firstWorksheet.Cells[i, 3].Value);
        //                }

        //                List<object> columnD = new List<object>();

        //                for (int i = 2; i< 56; i++)
        //                {
        //                    columnD.Add(firstWorksheet.Cells[i, 4].Value);
        //                }

        //                List<object> columnF = new List<object>();

        //                for (int i = 2; i< 56; i++)
        //                {
        //                    columnF.Add(firstWorksheet.Cells[i, 6].Value);
        //                }

        //                List<object> columnI = new List<object>();

        //                for (int i = 2; i< 56; i++)
        //                {
        //                    columnI.Add(firstWorksheet.Cells[i, 9].Value);
        //                }

        //                List<int> nullValues = new List<int>();

        //int k = 0;
        //                foreach (var item in columnB)
        //                {
        //                    k++;
        //                    if (item == null)
        //                    {
        //                        nullValues.Add(k-1);
        //                    }
        //                }

        //                for (int i = 0; i< 50; i++)
        //                {
        //                    if (!((nullValues[0] == i) || (nullValues[1] == i) || (nullValues[2] == i)))
        //                    {

        //                        using (System.IO.StreamWriter file =
        //                         new System.IO.StreamWriter(@"C:\Users\14168\Desktop\EmployeeMatrix.txt", true))
        //                        {
        //                            file.WriteLine(columnB[i].ToString()+ columnC[i].ToString() + DateTime.Parse(columnD[i].ToString()).Year+ columnF[i].ToString()+ columnI[i].ToString());
        //                        }
        //                        Console.WriteLine(" operator #: {0}, operator name: {1}, Hire date:{2},PayGrade:{3}, primary:{4}", columnB[i], columnC[i], DateTime.Parse(columnD[i].ToString()).Year, columnF[i], columnI[i]);

        //                    }

        //                }



        //                }


        //                Console.ReadKey();




        //            }



    }

    }

