using ExportToExcel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
//using Excel = Microsoft.Office.Interop.Excel;

namespace TaxesGovAzToExcelEVHF
{
    public class EVHF
    {
        public string IO { get; set; }
        public string Voen { get; set; }
        public string Ad { get; set; }
        public string Tip { get; set; }
        public string Veziyyet { get; set; }
        public string Tarix { get; set; }
        public string Seriya { get; set; }
        public string Nomre { get; set; }
        public string EsasQeyd { get; set; }
        public string ElaveQeyd { get; set; }
        public string EDVsiz { get; set; }
        public string EDV { get; set; }
        public string Hesab1C { get; set; }
        public string MVQeyd { get; set; }

        public override string ToString()
        {
            return $"{IO}-{Voen}-{Ad}-{Tip}-{Veziyyet}-{Tarix}-{Seriya}-{Nomre}-{EsasQeyd}-{ElaveQeyd}-{EDVsiz}-{EDV}-{Hesab1C}-{MVQeyd}";
        }

        public static List<EVHF> RZLoadEVHF(string link)
        {
            var RZEVHFList = new List<EVHF>();
            var RZEVHF = new EVHF();

/*
            // The HtmlWeb class is a utility class to get the HTML over HTTP
            HtmlWeb htmlWeb = new HtmlWeb();

            // Creates an HtmlDocument object from an URL
            HtmlDocument document = htmlWeb.Load(link);

            // Targets a specific node
            HtmlNode someNode = document.GetElementbyId("trback2");

            // If there is no node with that Id, someNode will be null
            if (someNode != null)
            {
                // Extracts all links within that node
                IEnumerable<HtmlNode> allLinks = someNode.Descendants("td");

                Console.WriteLine(allLinks.Count<HtmlNode>());
                // Outputs the href for external links
                foreach (HtmlNode linki in allLinks)
                {
                    Console.WriteLine(linki.InnerHtml);

                    // Checks whether the link contains an HREF attribute
                    //if (linki.Attributes.Contains("trback2"))
                    //{
                        // Simple check: if the href begins with "http://", prints it out
                        //if (linki.Attributes["trback2"].Value.StartsWith("http://"))
                    //        Console.WriteLine(linki.Attributes["trback2"].Value);
                    //}
                    //Console.WriteLine(linki);
                }
            }
*/

            //var temp = Path.GetTempFileName();
            //var tempFile = temp.Replace(Path.GetExtension(temp), ".html");
            //using (System.IO.StreamWriter sw = new System.IO.StreamWriter(tempFile))
            //{
            //    sw.Write("C:\\text.html");
            //}
            //Process.Start(new ProcessStartInfo(tempFile));

            // From File
            var doc1 = new HtmlDocument();
            Console.WriteLine(link);
            doc1.Load(link);
            

            // From String
            //var doc2 = new HtmlDocument();
            //doc2.LoadHtml(link);

            string tempDoc = doc1.ParsedText;
            string newTempDoc = tempDoc.Replace("ЖЏ", "Ə");
            newTempDoc = newTempDoc.Replace("Й™", "ə");
            newTempDoc = newTempDoc.Replace("Г–", "Ö");
            newTempDoc = newTempDoc.Replace("Г¶", "ö");
            newTempDoc = newTempDoc.Replace("Дћ", "Ğ");
            newTempDoc = newTempDoc.Replace("Дџ", "ğ");
            newTempDoc = newTempDoc.Replace("Д°", "İ");
            newTempDoc = newTempDoc.Replace("Д±", "ı"); 
            newTempDoc = newTempDoc.Replace("Гњ", "Ü");
            newTempDoc = newTempDoc.Replace("Гј", "ü");
            newTempDoc = newTempDoc.Replace("Г‡", "Ç"); 
            newTempDoc = newTempDoc.Replace("Г§", "ç");
            newTempDoc = newTempDoc.Replace("Ећ", "Ş");
            newTempDoc = newTempDoc.Replace("Еџ", "ş");
//            newTempDoc = newTempDoc.Replace("&nbsp;", "");

            newTempDoc = newTempDoc.Replace("<style>#trback{background-color:#dfe8f6;font-family : Tahoma;font-style : normal;font-size : 12px;font-weight : 100;}#trback2{background-color:#DFDFDF;font-family : Tahoma;font-style : normal;font-size : 12px;font-weight : 100;}#head{ font-family : Tahoma;font-style:normal;font-size : 14px;font-weight : 100;font:bold;text-align:center;color : #36428b;background-color:#a9c3ec}#qutu{border-left:1px solid #dfe8f6;border-bottom:1px solid #dfe8f6;border-right:1px solid #dfe8f6;border-top:1px solid #dfe8f6;}</style>", "");
            newTempDoc = newTempDoc.Replace("<HTML>", "");
            //newTempDoc = newTempDoc.Replace("<HEAD><meta http-equiv="Content - Type" content="text / html; charset = utf - 8" /><TITLE>VHF axtarışının nəticəsi</TITLE></HEAD>", "");
            newTempDoc = newTempDoc.Replace("<BODY>  <b> Axtarış şərtləri :<b>", "");
            //newTempDoc = newTempDoc.Replace("<i>Səhifə:Gələnlər, Tarix:21.05.2018-dən 21.05.2018-dək </i>", "");
            newTempDoc = newTempDoc.Replace("<br/>-----\n\n", "");



            // From Web
            //var url = "http://html-agility-pack.net/";
            //var web = new HtmlWeb();
            //var doc3 = web.Load(url);
            int j = 0, k = 0, count = 0;
            for (int i = 0; i < newTempDoc.Length; i++)
            {
                string tempDocx = "";
                for (; j < 6; j++)
                {
                    tempDocx += newTempDoc[(i+j) >= newTempDoc.Length-1 ? newTempDoc.Length-1 : (i+j)];
                }
//                Console.WriteLine(tempDocx);
                if (tempDocx == "&nbsp;")
                {
                    count++;
                    string Xvalue = "";
                    do
                    {
                        Xvalue += newTempDoc[(i + j + k) >= newTempDoc.Length-1 ? newTempDoc.Length-1 : (i + j + k)];
                        k++;
                    } while (newTempDoc[(i + j + k) >= newTempDoc.Length-1 ? newTempDoc.Length-1 : (i + j + k)] != '<');
                    k = 0;
                    //i = i + j + k + Xvalue.Length;
                    i = (i + j + k + Xvalue.Length) >= newTempDoc.Length-1 ? newTempDoc.Length-1 : (i + j + k + Xvalue.Length);
                    if (count == 1) RZEVHF.Voen = Xvalue;
                    if (count == 2) RZEVHF.Ad = Xvalue;
                    if (count == 3) RZEVHF.Tip = Xvalue;
                    if (count == 4) RZEVHF.Veziyyet = Xvalue;
                    if (count == 5) RZEVHF.Tarix = Xvalue;
                    if (count == 6) RZEVHF.Seriya = Xvalue;
                    if (count == 7) RZEVHF.Nomre = Xvalue;
                    if (count == 8) RZEVHF.EsasQeyd = Xvalue;
                    if (count == 9) RZEVHF.ElaveQeyd = Xvalue;
                    if (count == 10) RZEVHF.EDVsiz = Xvalue;
                    if (count == 11)
                    {
                        RZEVHF.EDV = Xvalue;
                        Console.WriteLine(RZEVHF.ToString());
                        count = 0;
                        RZEVHFList.Add(RZEVHF);
                    }
                    //public string IO { get; set; }
                    //public string Voen { get; set; }
                    //public string Ad { get; set; }
                    //public string Tip { get; set; }
                    //public string Veziyyet { get; set; }
                    //public string Tarix { get; set; }
                    //public string Seriya { get; set; }
                    //public string Nomre { get; set; }
                    //public string EsasQeyd { get; set; }
                    //public string ElaveQeyd { get; set; }
                    //public string EDVsiz { get; set; }
                    //public string EDV { get; set; }
                    //public string Hesab1C { get; set; }
                    //public string MVQeyd { get; set; }
                }
                j = 0;
            }
            return RZEVHFList;
        }
        public static void btnCreateExcel_Click(ref List<EVHF> EVHFs)
        {
//#if DEBUG
            //  We'll attempt to create our example .XLSX file in our "My Documents" folder
            string MyDocumentsPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string TargetFilename = System.IO.Path.Combine(MyDocumentsPath, "Sample.xlsx");
//#else
            // Prompt the user to enter a path/filename to save an example Excel file to
            //saveFileDialog1.FileName = "Sample.xlsx";
            //saveFileDialog1.Filter = "Excel 2007 files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //saveFileDialog1.FilterIndex = 1;
            //saveFileDialog1.RestoreDirectory = true;
            //saveFileDialog1.OverwritePrompt = false;

            ////  If the user hit Cancel, then abort!
            //if (saveFileDialog1.ShowDialog() != DialogResult.OK)
            //    return;

            //string TargetFilename = saveFileDialog1.FileName;
//#endif

            //  Step 1: Create a DataSet, and put some sample data in it
            DataSet ds = ExportToExcel(ref EVHFs);

            //  Step 2: Create the Excel file
            try
            {
                CreateExcelFile.CreateExcelDocument(ds, TargetFilename);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Couldn't create Excel file.\r\nException: " + ex.Message);
                return;
            }

            //  Step 3:  Let's open our new Excel file and shut down this application.
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo(TargetFilename);
            p.Start();

            //this.Close();
        }
        public static DataSet ExportToExcel(ref List<EVHF> EVHFs)
        {
            //  Create a sample DataSet, containing three DataTables.
            //  (Later, this will save to Excel as three Excel worksheets.)
            //
            DataSet ds = new DataSet();
            //  Create the first table of sample data
            DataTable dt1 = new DataTable("Drivers");

            dt1.Rows.Add();
            dt1.Rows.Add();

            dt1.Columns.Add("I/O", Type.GetType("System.String"));/*System.Decimal*/
            dt1.Columns.Add("VÖEN", Type.GetType("System.String"));
            dt1.Columns.Add("Adı", Type.GetType("System.String"));
            dt1.Columns.Add("Tipi", Type.GetType("System.String"));
            dt1.Columns.Add("Vəziyyəti", Type.GetType("System.String"));
            dt1.Columns.Add("VHF tarixi", Type.GetType("System.String"));
            dt1.Columns.Add("VHF Seriyası", Type.GetType("System.String"));
            dt1.Columns.Add("VHF nömrəsi", Type.GetType("System.String"));
            dt1.Columns.Add("Əsas qeyd", Type.GetType("System.String"));
            dt1.Columns.Add("Əlavə qeyd", Type.GetType("System.String"));
            dt1.Columns.Add("Malın ƏDV-siz ümumi dəyəri", Type.GetType("System.String"));
            dt1.Columns.Add("Malın ƏDV məbləği", Type.GetType("System.String"));

            //foreach (var i in EVHFs)
            //{
            //    dt1.Rows.Add(new object[] { "I", i.Voen, i.Ad, i.Tip, i.Veziyyet, i.Tarix, i.Seriya, i.Nomre, i.EsasQeyd, i.ElaveQeyd, i.EDVsiz, i.EDV});
            //}
            for (int i = 0; i < EVHFs.Count; i++)
            {
                dt1.Rows.Add(new object[] { "I", EVHFs[i].Voen,
                                                 EVHFs[i].Ad,
                                                 EVHFs[i].Tip,
                                                 EVHFs[i].Veziyyet,
                                                 EVHFs[i].Tarix,
                                                 EVHFs[i].Seriya,
                                                 EVHFs[i].Nomre,
                                                 EVHFs[i].EsasQeyd,
                                                 EVHFs[i].ElaveQeyd,
                                                 EVHFs[i].EDVsiz,
                                                 EVHFs[i].EDV });
            }
            ds.Tables.Add(dt1);
            return ds;
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Pleese inser Link");
            string link = "https://vroom.e-taxes.gov.az/index/index/" +
                "printServlet?tkn=MTcxMjU5MDMwMjIxNDMwNzA5ODQsMUhSUkIxWiwxLDE1Mjc1NzcwNDkxMDIsMDA3NDc1MTE=" +
                "&w=2" +
                "&v=" +
                "&fd=20180529000000" +
                "&td=20180529000000&s=" +
                "&n=" +
                "&sw=0" +
                "&r=1" +
                "&sv=1501069851";
            //link = Console.ReadLine();
            link = @"C:\text.html";
            
            List<EVHF> EVHFs = new List<EVHF>();
            EVHFs = EVHF.RZLoadEVHF(link);
            //EVHFs.Add();
            EVHF.btnCreateExcel_Click(ref EVHFs);
            Console.ReadKey();
        }
    }
}
