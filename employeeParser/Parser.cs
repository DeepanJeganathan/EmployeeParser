using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace employeeParser
{
    public class Parser
    {
        private string _filePath;
        private Dictionary<int, List<object>> _fullList;
        private int _startRow;
        private int _lastRow;
        private List<int> _columnNum;
        private List<string> _txt;
        public Parser(string filePath)
        {
            this._filePath = filePath;
        }

        public void CreateExcelPackage(string WorksheetName, List<int> ColumnNum, int startRow, int lastRow)
        {
            FileInfo fi = new FileInfo(_filePath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet Worksheet = excelPackage.Workbook.Worksheets[WorksheetName];

                var fullList = new Dictionary<int, List<object>>();
                foreach (var column in ColumnNum)
                {
                    
                    List<object> columnList = new List<object>();

                    for (int i = startRow; i < lastRow; i++)
                    {
                        columnList.Add(Worksheet.Cells[i,column].Value);
                    }

                    fullList.Add(column, columnList);

                   
                }

                this._fullList = fullList;
                this._startRow = startRow;
                this._lastRow = lastRow;
                this._columnNum = ColumnNum;
            }
          

        }

        public Dictionary<int,List<object>> GetDict()
        {
            return _fullList;

        }

        public void GetFormattedList()
        {
            var ans = _fullList;

            List<List<object>> ansList = new List<List<object>>();

            foreach (var item in ans)
            {
                ansList.Add(item.Value);
            }

            List<string> txt = new List<string>();

            for (int i = 0; i < (_lastRow - _startRow); i++)
            {
                var stringAns = "";
                for (int k = 0; k < _columnNum.Count; k++)
                {
                   
                    stringAns+= ansList[k][i];
                }
                txt.Add(stringAns);

                Console.WriteLine(stringAns);

                stringAns = "";

            }

            foreach (var item in txt)
            {
                Console.WriteLine(item);
            }

            this._txt = txt;


        }

        public void ExportToTxtFile(string TextFilePath)
        {
            GetFormattedList();
            foreach (var item in _txt)
            {
                using (System.IO.StreamWriter file =
                                                new System.IO.StreamWriter(TextFilePath, true))
                {
                    file.WriteLine(item);
                }

            }

        }


    }
}

 