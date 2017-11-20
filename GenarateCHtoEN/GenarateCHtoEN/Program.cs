using System;
using System.IO;
using System.Text;
using System.Net;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
namespace GenarateCHtoEN
{
    class Program
    {
        public static List<transDictionary> translateDictionary = new List<transDictionary>();
        static void Main(string[] args)
        {
            DataTable dt;
            //1.SET:translate chinese to english  excel file path 
            var excelpath = @"E:\转换工具\GenarateCHtoEN\GenarateCHtoEN\Excel\chToEnglish.xlsx";
            //read excel file
            dt = ExcelToDataTable(excelpath);
            foreach (DataRow row in dt.Rows)
            {
                translateDictionary.Add(new transDictionary
                {
                    name_ch = row.ItemArray[0].ToString(),
                    name_en = row.ItemArray[1].ToString()
                });

            }
            //2.SET:foreach file dictionary
            FindFile("F:/test");
            Console.WriteLine("Success: chinese to english translate completed. \n");
            Console.WriteLine("Press andy key to exit......");
            Console.ReadKey(true);
        }
        //创建html文件
        public static void createHtml(string fileName, string path)
        {
            path = System.IO.Path.Combine(path, fileName);
            Console.WriteLine("Path to my file: {0}\n", path);
            if (!System.IO.File.Exists(path))
            {
                using (System.IO.FileStream fs = System.IO.File.Create(path)) ;
            }
            else
            {
                Console.WriteLine("File \"{0}\" already exists.", fileName);
                return;
            }
        }

        //写入内容到html
        public static void writeToHtml(string fileName, string path, string content)
        {
            string fileAndName = path + "/" + fileName;
            using (FileStream fs = new FileStream(fileAndName, FileMode.Create))
            {
                using (StreamWriter w = new StreamWriter(fs, Encoding.UTF8))
                {
                    w.WriteLine(content);
                }
            }
        }

        //得到制定url的页面的code
        public static string getCode(string url)
        {
            string htmlCode;
            using (WebClient client = new WebClient())
            {
                htmlCode = client.DownloadString(url);
            }
            return htmlCode;
        }

        //制定excel转换为datatable
        public static DataTable ExcelToDataTable(string excelPath)
        {
            IWorkbook workbook;
            using (FileStream stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
            DataTable dt = new DataTable(sheet.SheetName);

            // write header row
            IRow headerRow = sheet.GetRow(0);
            foreach (ICell headerCell in headerRow)
            {
                dt.Columns.Add(headerCell.ToString());
            }

            // write the rest
            int rowIndex = 0;
            foreach (IRow row in sheet)
            {
                // skip header row
                if (rowIndex++ == 0) continue;
                DataRow dataRow = dt.NewRow();
                dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
                dt.Rows.Add(dataRow);
            }
            return dt;
        }

        //指定目录遍历
        public static void FindFile(string dirPath)
        {

            DirectoryInfo Dir = new DirectoryInfo(dirPath);
            foreach (DirectoryInfo d in Dir.GetDirectories())
            {
                FindFile(Dir + @"\" + d.ToString());
            }
            foreach (FileInfo f in Dir.GetFiles("*.html").Union(Dir.GetFiles("*.js")).ToArray()) //查找文件
            {
                string fullName = f.FullName;
                string getPath = f.DirectoryName;
                string ChName = f.Name;
                string name_en = "";
                string filterName = "_en.html";
                //filter english file
                if (ChName.Contains(filterName)) {
                    break;
                }
                //生成英文文件名
                foreach (char c in ChName)
                {
                    if (c == '.')
                    {
                        name_en += "_en.";
                        continue;
                    }
                    name_en += c;
                }
                //读取当前文件代码
                string content = getCode(@" " + fullName + " ");
                //替换对应的中文转换为英文
                foreach (var item in translateDictionary)
                {
                    var nameCH = item.name_ch;
                    var nameEN = item.name_en;
                    content = cshtmlRepalce(content, nameCH, nameEN);
                }
                //创建HTML文件
                createHtml(name_en, getPath);
                //写入替换后的代码
                writeToHtml(name_en, getPath, content);

            }
        }
        /// <summary>
        /// 根据正则替换chhtml中指定字符串为指定内容
        /// </summary>
        /// <param name="input">cshtml内容</param>
        /// <param name="searchText">搜索的内容</param>
        /// <param name="repalceText">替换的内容</param>
        /// <returns>替换后的结果</returns>
        public static string cshtmlRepalce(string input, string searchText, string repalceText)
        {

            searchText = Regex.Replace(searchText, @"\[|\]|\.|\(|\)", @"\$0");
            Regex reg = new Regex(@"(?<![\u4e00-\u9fa5])" + searchText + @"(?![\u4e00-\u9fa5])");
            return reg.Replace(input, repalceText);
        }
    }
}
