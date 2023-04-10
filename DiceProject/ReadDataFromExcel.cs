using System; //using the system library in my project, giving access to classes and functions like Console and WriteLine like the ones below.
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace DiceProject
{
    public class ReadDataFromExcel
    {
        private Stream InFile;
        private IWorkbook Book;
        private ISheet Sheet;
        private Dictionary<string, int> Headers;

        public void DataFile(string filepath, string sheetName)
        {
            // Open a file stream to the excel file
            InFile = new FileStream(filepath, FileMode.Open);

            // Open the Workbook and Worksheet for the Excel file
            Book = WorkbookFactory.Create(InFile);
            Sheet = Book.GetSheet(sheetName);

            // Keep track of the column headers
            Headers = new Dictionary<string, int>();
            IRow headerRow = Sheet.GetRow(0);
            for (int col = 0; col < 2; col++)
            {
                string curr = headerRow.GetCell(col).StringCellValue;
                Headers.Add(curr, col);
            }
        }
        public void DataFile2(string filepath, string sheetName)
        {
            int Count = 0;

            // Open a file stream to the excel file
            InFile = new FileStream(filepath, FileMode.Open);

            // Open the Workbook and Worksheet for the Excel file
            Book = WorkbookFactory.Create(InFile);
            Sheet = Book.GetSheet(sheetName);

            // Keep track of the column headers
            Headers = new Dictionary<string, int>();
            
            IRow headerRow = Sheet.GetRow(0);
            S3 = headerRow.GetCell(num).ToString();
            while (S3 != "" || S3 ==null)
            {
              
                if (headerRow.GetCell(num) != null || Convert.ToString(headerRow.GetCell(num)) != "" )
                {
                    S3 = headerRow.GetCell(num).ToString();
                }
                    
                
                num++;
                Count = Count + 1;
            }

            num = num - 1;

            String curr2 = headerRow.GetCell(Count-2).DateCellValue.ToString("MM-dd-yy");
            DateTime dateTime = DateTime.Now.Date;
            DateTime d2 = dateTime;
            var dateString2 = DateTime.Now.ToString("MM-dd-yy");
            if (curr2 == dateString2)
            {
                num = num - 1;
            }

        }
        public DateTime? StrToDate(string val)
        {
            DateTime? dt = string.IsNullOrEmpty(val)
                ? (DateTime?)null
                : DateTime.ParseExact(val, "MM-dd-yy", null);
            return dt;
        }


        /// Returns the data currently in the specified cell
        public String GetDataFromColumn(string columnName, int rowNum)
        {
            int colNum;
            ICell cell;

            // Return a null if the row number is invalid
            if (rowNum > Sheet.LastRowNum)
            {
                return null;
            }

            // Get the column number from the column name
            try
            {
                colNum = Headers[columnName];
            }
            catch (KeyNotFoundException)
            {
                // Return a null if the key is invalid
                Console.WriteLine("No data found in Excel file");
                return null;
            }

            // Return the data from the specified cell
            cell = Sheet.GetRow(rowNum).GetCell(colNum);

            // Return data based on Cell type
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Blank:
                    return "";
                case CellType.Formula:
                    IFormulaEvaluator eval = Book.GetCreationHelper().CreateFormulaEvaluator();

                    switch (eval.EvaluateFormulaCell(cell))
                    {
                        case CellType.String:
                            return cell.StringCellValue;
                        case CellType.Numeric:
                            return cell.NumericCellValue.ToString();
                        case CellType.Boolean:
                            return cell.BooleanCellValue.ToString();
                        case CellType.Blank:
                            return "";
                    }
                    break;
            }

            return null;
        }
        public String GetDataFromColumn2(int colNum, int rowNum)
        {       
            ICell cell;

            // Return a null if the row number is invalid
            if (rowNum > Sheet.LastRowNum)
            {
                return null;
            }

            cell = Sheet.GetRow(rowNum).GetCell(colNum);
                    return cell.StringCellValue;

        }
        public int RowCount
        {
            get
            {
                return Sheet.LastRowNum;
            }
        }

        /// Closes the Excel file
        public void Close()
        {
            InFile.Close();
        }
        private int _num;
        private string _s3;

        public int num
        {
            get
            {
                return _num;
            }
            set
            {
                _num = value;
            }
        }
        public String S3
        {
            get
            {
                return _s3;
            }
            set
            {
                _s3 = value;
            }
        }
        private int _num2;
        public int num2
        {
            get
            {
                return _num2;
            }
            set
            {
                _num2 = value;
            }
        }
    }
}
