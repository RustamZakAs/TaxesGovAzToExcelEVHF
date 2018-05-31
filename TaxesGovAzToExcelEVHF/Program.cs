using ExportToExcel;
using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;
//using System.IO.PathTooLongException;
//using Excel = Microsoft.Office.Interop.Excel;

namespace TaxesGovAzToExcelEVHF
{
    public class EVHF
    {
        public EVHF() {}
        public EVHF(string iO, string voen, string ad, string tip, string veziyyet, string tarix, string seriya, string nomre, string esasQeyd, string elaveQeyd, string eDVsiz, string eDV, string hesab1C, string mVQeyd)
        {
            IO = iO;
            Voen = voen;
            Ad = ad;
            Tip = tip;
            Veziyyet = veziyyet;
            Tarix = tarix;
            Seriya = seriya;
            Nomre = nomre;
            EsasQeyd = esasQeyd;
            ElaveQeyd = elaveQeyd;
            EDVsiz = eDVsiz;
            EDV = eDV;
            Hesab1C = hesab1C;
            MVQeyd = mVQeyd;
        }

        public EVHF(EVHF eVHF)
        {
            IO = eVHF.IO;
            Voen = eVHF.Voen;
            Ad = eVHF.Ad;
            Tip = eVHF.Tip;
            Veziyyet = eVHF.Veziyyet;
            Tarix = eVHF.Tarix;
            Seriya = eVHF.Seriya;
            Nomre = eVHF.Nomre;
            EsasQeyd = eVHF.EsasQeyd;
            ElaveQeyd = eVHF.ElaveQeyd;
            EDVsiz = eVHF.EDVsiz;
            EDV = eVHF.EDV;
            Hesab1C = eVHF.Hesab1C;
            MVQeyd = eVHF.MVQeyd;
        }

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

        public static void RZLoadEVHF(ref List<EVHF> EVHFList, string[] link)
        {
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

            // From File
            var doc1 = new HtmlDocument();
            //< runtime >
            //< AppContextSwitchOverrides value = "Switch.System.IO.UseLegacyPathHandling=false;Switch.System.IO.BlockLongPaths=false" />
            //</ runtime >
            for (int i = 0; i < link.Length; i++)
            {
                Console.WriteLine(link.Length); 
                Console.WriteLine(link[i]);
            }

            //**********TEST ERROR BEGIN**************
            //string reallyLongDirectory = @"C:\Test\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            //reallyLongDirectory = reallyLongDirectory + @"\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            //reallyLongDirectory = reallyLongDirectory + @"\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";

            //Console.WriteLine($"Creating a directory that is {reallyLongDirectory.Length} characters long");
            //Directory.CreateDirectory(reallyLongDirectory);
            //**********TEST ERROR END**************

            var temp = Path.GetTempFileName();
            var tempFile = temp.Replace(Path.GetExtension(temp), ".html");
            for (int i = 0; i < link.Length; i++)
            {
                try
                {
                    using (StreamWriter sw = new StreamWriter(tempFile))
                    {
                        for (int j = 0; j < link.Length; j++)
                        {
                            sw.Write($"{link[j]}<br>");
                        }
                    }
                }
                catch (Exception e) 
                {
                    Console.WriteLine(e.Message);
                }

                try
                {
                    // Open the text file using a stream reader.
                    //using (StreamReader sr = new StreamReader(link)) //link = "TestFile.txt"
                    //{
                    //    // From Web
                    //    //var url = @"http://html-agility-pack.net/";
                    //    //var web = new HtmlWeb();
                    //    //var doc3 = web.Load(url);
                    //
                    //    // From String
                    //    //var doc2 = new HtmlDocument();
                    //    //doc2.LoadHtml(link);
                    //
                    //    // Read the stream to a string, and write the string to the console.
                    //    String line = sr.ReadToEnd();
                    //    Console.WriteLine(line);
                    //}
                    //---doc1.Load(link[i]);
                }
                catch (Exception e) //******************ERROR LEN************** 260 norm, my link 202
                {
                    Console.WriteLine("The file could not be read:");
                    Console.WriteLine(e.Message);
                }

                //---string tempDoc = doc1.ParsedText;
                //---string newTempDoc = tempDoc.Replace("ЖЏ", "Ə");
                //---newTempDoc = newTempDoc.Replace("Й™", "ə");
                //---newTempDoc = newTempDoc.Replace("Г–", "Ö");
                //---newTempDoc = newTempDoc.Replace("Г¶", "ö");
                //---newTempDoc = newTempDoc.Replace("Дћ", "Ğ");
                //---newTempDoc = newTempDoc.Replace("Дџ", "ğ");
                //---newTempDoc = newTempDoc.Replace("Д°", "İ");
                //---newTempDoc = newTempDoc.Replace("Д±", "ı"); 
                //---newTempDoc = newTempDoc.Replace("Гњ", "Ü");
                //---newTempDoc = newTempDoc.Replace("Гј", "ü");
                //---newTempDoc = newTempDoc.Replace("Г‡", "Ç"); 
                //---newTempDoc = newTempDoc.Replace("Г§", "ç");
                //---newTempDoc = newTempDoc.Replace("Ећ", "Ş");
                //---newTempDoc = newTempDoc.Replace("Еџ", "ş");
                //newTempDoc = newTempDoc.Replace("&nbsp;", "");

                //---newTempDoc = newTempDoc.Replace("<style>#trback{background-color:#dfe8f6;font-family : Tahoma;font-style : normal;font-size : 12px;font-weight : 100;}#trback2{background-color:#DFDFDF;font-family : Tahoma;font-style : normal;font-size : 12px;font-weight : 100;}#head{ font-family : Tahoma;font-style:normal;font-size : 14px;font-weight : 100;font:bold;text-align:center;color : #36428b;background-color:#a9c3ec}#qutu{border-left:1px solid #dfe8f6;border-bottom:1px solid #dfe8f6;border-right:1px solid #dfe8f6;border-top:1px solid #dfe8f6;}</style>", "");
                //---newTempDoc = newTempDoc.Replace("<HTML>", "");
                //---newTempDoc = newTempDoc.Replace("<HEAD><meta http-equiv=\"Content - Type\" content=\"text / html; charset = utf - 8\" /><TITLE>VHF axtarışının nəticəsi</TITLE></HEAD>", "");
                //---newTempDoc = newTempDoc.Replace("<BODY>  <b> Axtarış şərtləri :<b>", "");
                //---newTempDoc = newTempDoc.Replace("<i>Səhifə:Gələnlər, Tarix:", "");
                //---newTempDoc = newTempDoc.Replace("<br/>-----\n\n", "");

                //---EVHFList.AddRange(StringToListEVHF(newTempDoc));
            }
            Process.Start(new ProcessStartInfo(tempFile));
        }
        public static List<EVHF> StringToListEVHF(string str)
        {
            //List<EVHF> RZEVHFList = new List<EVHF>();

            var RZEVHFList = new List<EVHF>();
            var RZEVHF = new EVHF();
            //string[] RZEVHFstring = new string[14];

            int j = 0, k = 0, count = 0;
            for (int i = 0; i < str.Length; i++)
            {
                string tempDocx = "";
                for (; j < MainEVHF.TextForBegin.Length; j++)
                {
                    tempDocx += str[(i + j) >= str.Length - 1 ?
                        str.Length - 1 : (i + j)];
                }
                if (tempDocx == MainEVHF.TextForBegin)
                {
                    count++;
                    string Xvalue = "";
                    int tempIndex = 0;
                    do
                    {
                        tempIndex = (i + j + k) >= str.Length - 1 ?
                            str.Length - 1 : (i + j + k);
                        Xvalue += str[tempIndex];
                        k++;
                    } while (str[tempIndex] != '<');
                    Xvalue = Xvalue.Replace("<", "");
                    tempIndex = 0;
                    k = 0;
                    i = (i + j + k + Xvalue.Length) >= str.Length - 1 ?
                        str.Length - 1 : (i + j + k + Xvalue.Length);
                    if (count == 1)
                    {
                        RZEVHF/*[0]*/.IO = MainEVHF.EVHFIO;
                        RZEVHF/*[1]*/.Voen = Xvalue;
                    }
                    if (count == 2) RZEVHF/*[2]*/.Ad = Xvalue;
                    if (count == 3) RZEVHF/*[3]*/.Tip = Xvalue;
                    if (count == 4) RZEVHF/*[4]*/.Veziyyet = Xvalue;
                    if (count == 5) RZEVHF/*[5]*/.Tarix = Xvalue;
                    if (count == 6) RZEVHF/*[6]*/.Seriya = Xvalue;
                    if (count == 7) RZEVHF/*[7]*/.Nomre = Xvalue;
                    if (count == 8) RZEVHF/*[8]*/.EsasQeyd = Xvalue;
                    if (count == 9) RZEVHF/*[9]*/.ElaveQeyd = Xvalue;
                    if (count == 10)
                    {
                        //Xvalue = Xvalue.Replace(".", ",");
                        //RZEVHF.EDVsiz = decimal.Parse(Xvalue);
                        RZEVHF/*[10]*/.EDVsiz = Xvalue;
                    }
                    if (count == 11)
                    {
                        //Xvalue = Xvalue.Replace(".", ",");
                        //RZEVHF.EDV = decimal.Parse(Xvalue);
                        RZEVHF/*[11]*/.EDV = Xvalue;
                        RZEVHF/*[12]*/.Hesab1C = "531.1";
                        RZEVHF/*[13]*/.MVQeyd = "";
                        //Console.WriteLine(RZEVHF.ToString());
                        RZEVHFList.Add(new EVHF(RZEVHF));
                        //RZEVHFList.Add(new EVHF(RZEVHF[0], 
                        //    RZEVHF[1], 
                        //    RZEVHF[2], 
                        //    RZEVHF[3], 
                        //    RZEVHF[4], 
                        //    RZEVHF[5],
                        //    RZEVHF[6],
                        //    RZEVHF[7],
                        //    RZEVHF[8],
                        //    RZEVHF[9],
                        //    RZEVHF[10],
                        //    RZEVHF[11],
                        //    RZEVHF[12],
                        //    RZEVHF[13]));//**********ERROR***********
                        count = 0;
                    }
                }
                j = 0;
            }
            return RZEVHFList;
        }
        public static void CreateExcel(ref List<EVHF> EVHFs)
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
            Process p = new Process
            {
                StartInfo = new ProcessStartInfo(TargetFilename)
            };
            p.Start();

            //this.Close();
        }
        public static DataSet ExportToExcel(ref List<EVHF> EVHFs)
        {
            //  Create a sample DataSet, containing three DataTables.
            //  (Later, this will save to Excel as three Excel worksheets.)
            DataSet ds0 = new DataSet();
            //  Create the first table of sample data
            DataTable dt0 = new DataTable("EVHF");
            //dt0.Rows.Add(new object[] { "EVHF siyahısı" });
            //dt0.Rows.Add(new object[] { });

            dt0.Columns.Add("I/O", Type.GetType("System.String"));/*System.Decimal*/
            dt0.Columns.Add("VÖEN", Type.GetType("System.String"));
            dt0.Columns.Add("Adı", Type.GetType("System.String"));
            dt0.Columns.Add("Tipi", Type.GetType("System.String"));
            dt0.Columns.Add("Vəziyyəti", Type.GetType("System.String"));
            dt0.Columns.Add("VHF tarixi", Type.GetType("System.String"));
            dt0.Columns.Add("VHF Seriyası", Type.GetType("System.String"));
            dt0.Columns.Add("VHF nömrəsi", Type.GetType("System.String"));
            dt0.Columns.Add("Əsas qeyd", Type.GetType("System.String"));
            dt0.Columns.Add("Əlavə qeyd", Type.GetType("System.String"));
            dt0.Columns.Add("Malın ƏDV-siz ümumi dəyəri", Type.GetType("System.String"));
            dt0.Columns.Add("Malın ƏDV məbləği", Type.GetType("System.String"));
            dt0.Columns.Add("1C", Type.GetType("System.String"));
            dt0.Columns.Add("Malverən qeyd", Type.GetType("System.String"));

            foreach (var i in EVHFs)
            {
                dt0.Rows.Add(new object[] { "I", i.Voen, i.Ad, i.Tip, i.Veziyyet, i.Tarix, i.Seriya, i.Nomre, i.EsasQeyd, i.ElaveQeyd, i.EDVsiz, i.EDV, i.Hesab1C, i.MVQeyd });
            }
            //for (int i = 0; i < EVHFs.Count; i++)
            //{
            //    dt0.Rows.Add(new object[] { "I", EVHFs[i].Voen,
            //                                     EVHFs[i].Ad,
            //                                     EVHFs[i].Tip,
            //                                     EVHFs[i].Veziyyet,
            //                                     EVHFs[i].Tarix,
            //                                     EVHFs[i].Seriya,
            //                                     EVHFs[i].Nomre,
            //                                     EVHFs[i].EsasQeyd,
            //                                     EVHFs[i].ElaveQeyd,
            //                                     EVHFs[i].EDVsiz,
            //                                     EVHFs[i].EDV,
            //                                     EVHFs[i].Hesab1C,
            //                                     EVHFs[i].MVQeyd});
            //}
            ds0.Tables.Add(dt0);
            return ds0;
        }
    }
    class MainEVHF
    {
        //*****************************************
        private static string myIO;
        public static string EVHFIO
        {
            get { return myIO = "I"; }
            set { myIO = value; }
        }
        //*****************************************
        private static string myTextForBegin;
        public static string TextForBegin
        {
            get { return myTextForBegin = "&nbsp;"; }
            set { myTextForBegin = value; }
        }
        //*****************************************
        public static string EVHFsLink { get; set; }
        //*****************************************
        private static string myEVHFsVoen;
        public static string EVHFsVOEN
        {
            get { return myEVHFsVoen = "1501069851"; }
            set { myEVHFsVoen = value; }
        }
        //*****************************************
        public static string EVHFIlkTarix { get; set; }
        //*****************************************
        public static string EVHFSonTarix { get; set; }
        //*****************************************
        public static void MainMenyu ()
        {
            string link;
            bool tokenExsist = false;
            do
            {
                Console.Clear();
                Console.WriteLine("Pleese inser Link");
                link = Console.ReadLine();
                if (CopyToken(link).Length > 0) tokenExsist = true;
            } while (!tokenExsist);
            //link = //@"https://vroom.e-taxes.gov.az/index/index/" +
            //    "printServlet?tkn=MTcxMjU5MDMwMjIxNDMwNzA5ODQsMUhSUkIxWiwxLDE1Mjc1NzcwNDkxMDIsMDA3NDc1MTE=" +
            //    "&w=2" +
            //    "&v=" +
            //    "&fd=20180529000000" +
            //    "&td=20180529000000&s=" +
            //    "&n=" +
            //    "&sw=0" +
            //    "&r=1" +
            //    "&sv=1501069851";

            Console.WriteLine("Ilk tarixi daxil edin: YYYYMMDD");
            EVHFIlkTarix = Console.ReadLine();
            Console.WriteLine("Son tarixi daxil edin: YYYYMMDD");
            EVHFSonTarix = Console.ReadLine();
            Console.WriteLine($"VOEN: {CopyVoen(link)}");

            //link = @"https://vroom.e-taxes.gov.az/index/index/" +
            //        @"printServlet?tkn=" + CopyToken(link) + @"==" +
            //        @"&w=2" +
            //        @"&v=" +
            //        @"&fd=" + EVHFIlkTarix + @"000000" +
            //        @"&td=" + EVHFSonTarix + @"000000" +
            //        @"&s=" +
            //        @"&n=" +
            //        @"&sw=0" +
            //        @"&r=1" +
            //        @"&sv=" + EVHFsVOEN;
            //link = @"C:\text.html";
            EVHFsLink = link;
        }
        private static string CopyToken(string link)
        {
            string XToken = "";
            for (int i = 0; i < link.Length; i++)
            {
                string Xtemp = "";
                if (i < link.Length - 3)
                {
                    Xtemp += link[i + 0];
                    Xtemp += link[i + 1];
                    if (Xtemp == "t=")
                    {
                        int x = 0, xlen = 0;
                        do
                        {
                            xlen = i + Xtemp.Length + x++;
                            if (link[xlen] == '=' || link[xlen] == '&') break;
                            if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                        } while (true);
                        //XToken = XToken.Remove(XToken.Length - 1, 1);
                    }
                }
            }
            if (XToken.Length == 0)
            {
                for (int i = 0; i < link.Length; i++)
                {
                    string Xtemp = "";
                    if (i < link.Length - 3)
                    {
                        Xtemp += link[i + 0];
                        Xtemp += link[i + 1];
                        Xtemp += link[i + 2];
                        Xtemp += link[i + 3];
                        if (Xtemp == "tkn=")
                        {
                            int x = 0, xlen = 0;
                            do
                            {
                                xlen = i + Xtemp.Length + x++;
                                if (link[xlen] == '=' || link[xlen] == '&') break;
                                if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                            } while (true);
                            //XToken = XToken.Remove(XToken.Length - 1, 1);
                        }
                    }
                }
            }
            //Console.WriteLine(XToken);
            return XToken;
        }
        private static string CopyVoen(string link)
        {
            string XToken = "";
            for (int i = 0; i < link.Length; i++)
            {
                string Xtemp = "";
                if (i < link.Length - 3)
                {
                    Xtemp += link[i + 0];
                    Xtemp += link[i + 1];
                    if (Xtemp == "v=")
                    {
                        int x = 0, xlen = 0;
                        do
                        {
                            xlen = i + Xtemp.Length + x++;
                            if (link[xlen] == '=' || link[xlen] == '&') break;
                            if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                        } while (true);
                        //XToken = XToken.Remove(XToken.Length - 1, 1);
                    }
                }
            }
            if (XToken.Length == 0)
            {
                for (int i = 0; i < link.Length; i++)
                {
                    string Xtemp = "";
                    if (i < link.Length - 3)
                    {
                        Xtemp += link[i + 0];
                        Xtemp += link[i + 1];
                        Xtemp += link[i + 2];
                        Xtemp += link[i + 3];
                        if (Xtemp == "&sv=")
                        {
                            int x = 0, xlen = 0;
                            do
                            {
                                xlen = i + Xtemp.Length + x++;
                                if (link[xlen] == '=' || link[xlen] == '&') break;
                                if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                            } while (true);
                            //XToken = XToken.Remove(XToken.Length - 1, 1);
                        }
                    }
                }
            }
            //Console.WriteLine(XToken);
            return XToken;
        }
        private static string[] CreateLinkArray (string link, string beginDate, string endDate)
        {
            string xyear = "", xmonth = "", xday = "";
            xyear += beginDate[0];
            xyear += beginDate[1];
            xyear += beginDate[2];
            xyear += beginDate[3];
            xmonth += beginDate[4];
            xmonth += beginDate[5];
            xday += beginDate[6];
            xday += beginDate[7];

            DateTime beginDateTime = new DateTime(int.Parse(xyear), int.Parse(xmonth), int.Parse(xday));

            xyear = ""; xmonth = ""; xday = "";
            xyear += endDate[0];
            xyear += endDate[1];
            xyear += endDate[2];
            xyear += endDate[3];
            xmonth += endDate[4];
            xmonth += endDate[5];
            xday += endDate[6];
            xday += endDate[7];

            DateTime endDateTime = new DateTime(int.Parse(xyear), int.Parse(xmonth), int.Parse(xday));

            TimeSpan difference = endDateTime.Date - beginDateTime.Date;
            int days = (int)difference.TotalDays + 1;
            //Console.WriteLine(days);

            string[] linkArray = new string[days];
            
            DateTime tempDateTime;

            for (int i = 0; i < days; i++)
            {
                string strDate = "";
                tempDateTime = beginDateTime.AddDays(i);
                strDate = strDate + tempDateTime.Year.ToString();
                strDate += tempDateTime.Month.ToString().Length == 1 ? $"0{tempDateTime.Month.ToString()}" : $"{tempDateTime.Month.ToString()}";
                strDate += tempDateTime.Day.ToString().Length == 1 ? $"0{tempDateTime.Day.ToString()}" : $"{tempDateTime.Day.ToString()}";

                linkArray[i] = @"https://vroom.e-taxes.gov.az/index/index/" +
                    @"printServlet?tkn=" + CopyToken(link) + @"==" +
                    @"&w=2" +
                    @"&v=" +
                    @"&fd=" + strDate + @"000000" +
                    @"&td=" + strDate + @"000000" +
                    @"&s=" +
                    @"&n=" +
                    @"&sw=0" +
                    @"&r=1" +
                    @"&sv=" + EVHFsVOEN;
                //linkArray[i] = $"C:\\text{i}.html";
            }
            return linkArray;
        }
        public static void Main(string[] args)
        {
            MainMenyu();
            string[] linrArray = CreateLinkArray(EVHFsLink, EVHFIlkTarix, EVHFSonTarix);

            List<EVHF> EVHFlist = new List<EVHF>();
            EVHF.RZLoadEVHF(ref EVHFlist, linrArray);
            EVHF.CreateExcel(ref EVHFlist);
            Console.ReadKey();
        }
    }
}


/*
Process process = new Process();
process.StartInfo.FileName = "ipconfig.exe";        
process.StartInfo.UseShellExecute = false;
process.StartInfo.RedirectStandardOutput = true;        
process.Start();

// Synchronously read the standard output of the spawned process. 
StreamReader reader = process.StandardOutput;
string output = reader.ReadToEnd();

// Write the redirected output to this application's window.
Console.WriteLine(output);

process.WaitForExit();
process.Close();

Console.WriteLine("\n\nPress any key to exit.");
Console.ReadLine();
*/