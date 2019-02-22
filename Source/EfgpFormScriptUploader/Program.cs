using System;
using System.IO;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace EfgpFormScriptUploader {
    class Program {
        private static SqlConnection conn;
        private static ChromeDriver browser;
        private static string scriptCode;
        private static string databaseName;
        private static string hostName;
        private static string connString;
        private static ChromeDriverService driver;
        private static ChromeOptions options;
        private static string scriptPath;
        private static string formId;
        private static DataTable columnInformation;
        private static string oid;
        private static string formDesignerAjaxUrl;
        private static IJavaScriptExecutor console;

        static int Main(string[] args) {
            try {
                if (args.Length != 4) throw new Exception("必須有4個參數[表單Script檔位置],[BPM資料庫主機名稱],[BPM資料庫名稱],[BPM AP主機位址]");

                hostName = args[1];
                databaseName = args[2];
                connString = string.Format(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString, databaseName, hostName);

                driver = ChromeDriverService.CreateDefaultService();
                driver.HideCommandPromptWindow = true;
                options = new ChromeOptions();
                options.AddArguments("headless");
                browser = new ChromeDriver(driver, options);
                conn = new SqlConnection(connString);

                scriptPath = args[0];
                formId = Path.GetFileNameWithoutExtension(scriptPath);

                conn.Open();
                columnInformation = GetTableContent(conn, String.Format(@"SELECT OID FROM FormDefinition WHERE id = '{0}' AND publicationStatus = 'RELEASED';", formId));
                if (columnInformation.Rows.Count != 1) throw new Exception("無法取得表單,ID" + "[" + formId + "]資訊");

                oid = columnInformation.Rows[0][0].ToString();
                scriptCode = File.ReadAllText(scriptPath, Encoding.UTF8);

                formDesignerAjaxUrl = string.Format(ConfigurationManager.AppSettings["FormDesignerAjaxUrl"], args[3]);
                browser.Navigate().GoToUrl(formDesignerAjaxUrl);
                browser.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10000);

                console = (IJavaScriptExecutor)browser;
                console.ExecuteScript("DWREngine._execute(formDesignerAjax._path, 'formDesignerAjax', 'updateFormScript', '" + oid + "', '" + scriptCode.Replace("\\", "\\\\").Replace("'", "\\'").Replace(Environment.NewLine, "\\n") + "\\n\\n\\n', null );");

                Console.WriteLine("上傳完成");

                return 0;
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);

                return 1;
            }
            finally {
                if (conn != null) conn.Close();
                if (browser != null) browser.Quit();
            }
        }

        /// <summary>
        /// 取得Table內容
        /// </summary>
        /// <param name="conn">SQL連線</param>
        /// <param name="sqlSyntax">SQL語法</param>
        /// <returns>回傳結果</returns>
        static DataTable GetTableContent(SqlConnection conn, string sqlSyntax) {
            var dt = new DataTable();
            var command = new SqlCommand(sqlSyntax, conn);
            var reader = command.ExecuteReader();
            dt.Load(reader);

            return dt;
        }
    }
}