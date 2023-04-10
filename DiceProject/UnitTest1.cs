namespace DiceProject
{
    public class Tests
    {
        WebDriver driver;
        ReadDataFromExcel ReadDataFromExcel = new ReadDataFromExcel();
        String Link = "https:www.dice.com";
        String s;
        WorkBook WorkBook;
        String sheet;
        String outputfilePath = "C:\\Users\\Tenzi\\source\\repos\\DiceProject\\";
        DateTime dateTime = DateTime.Now;
        String d;
        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
        }

        [Test]
        public void Test1()
        {



            ReadDataFromExcel.DataFile("C:\\Users\\Tenzi\\source\\repos\\DiceProject\\InputFile.xlsx", "sheet data");
            if (File.Exists(outputfilePath + "search.xlsx"))
            {

                ReadDataFromExcel.DataFile2(outputfilePath + "search.xlsx", "sheet data");
                excelHeader2();
                for (int i = 1; i < 40; i++)
                {
                    driver.Navigate().GoToUrl(Link);
                    search_keyword(ReadDataFromExcel.GetDataFromColumn2(0, i));
                    excelwrite2();
                }
                excelfooter_save2();

                ReadDataFromExcel.Close();
            }
            else
            {
                excelHeader();
                for (int i = 1; i < 40; i++)
                {
                    driver.Navigate().GoToUrl(Link);
                    search_keyword(ReadDataFromExcel.GetDataFromColumn("KeyWords", i));

                    excelwrite(ReadDataFromExcel.GetDataFromColumn("KeyWords", i));
                }
                ReadDataFromExcel.Close();
                excelfooter_save();

            }
            driver.Close();
            Assert.Pass();
        }

        public void search_keyword(String value)
        {
            try
            {
                driver.FindElement(By.Id("typeaheadInput")).SendKeys(value);
                // Thread.Sleep(2000);
                driver.FindElement(By.Id("submitSearch-button")).Click();
                Thread.Sleep(2000);
                s = driver.FindElement(By.Id("totalJobCount")).Text;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);


            }
        }

        public void excelwrite(String KeyWords)
        {
            var sheet = WorkBook.GetWorkSheet("sheet data");
            int i = 2;

            sheet["A" + (i + num)].Value = KeyWords;
            sheet["B" + (i + num)].Value = s;

            s = "";
            num++;

        }
        public void excelwrite2()
        {
            WorkBook workbook = WorkBook.Load(outputfilePath + "_search.xlsx");
            WorkSheet sheet = workbook.DefaultWorkSheet;
            int i = 2;

            sheet[(Convert.ToChar(65 + ReadDataFromExcel.num)).ToString() + (i + num)].Value = s;

            s = "";
            num++;
            sheet.SaveAs(outputfilePath + "search.xlsx");
        }

        public void excelHeader()
        {

            WorkBook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = WorkBook.CreateWorkSheet("sheet data");
            DateTime dateTime = DateTime.Now.Date;
            d = dateTime.ToString("MM-dd-yy");

            sheet["A1"].Value = "KeyWords";
            sheet["B1"].Value = d;
            sheet["A1:D1"].Style.Font.Bold = true;

        }
        public void excelHeader2()
        {
            WorkBook workbook = WorkBook.Load(outputfilePath + "search.xlsx");
            WorkSheet sheet = workbook.DefaultWorkSheet;
            sheet.ToDataTable(true);

            DateTime dateTime = DateTime.Now.Date;
            d = dateTime.ToString("MM-dd-yy");
            string asciichar = (Convert.ToChar(65 + ReadDataFromExcel.num)).ToString();

            sheet[asciichar + 1].Value = d;
            sheet[asciichar + 1].Style.Font.Bold = true;
            sheet.SaveAs(outputfilePath + "search.xlsx");
        }

        public void excelfooter_save()
        {
            try
            {

                WorkBook.SaveAs(outputfilePath + "search.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);


            }
        }
        public void excelfooter_save2()
        {
            try
            {
                WorkBook workbook = WorkBook.Load(outputfilePath + "search.xlsx");
                WorkSheet sheet = workbook.DefaultWorkSheet;

                sheet.SaveAs(outputfilePath + "search.xlsx");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);


            }
        }
        private int _num;
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


    }
}