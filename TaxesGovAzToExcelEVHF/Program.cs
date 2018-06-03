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
    public class EVHF : IComparable//, IEnumerator
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

            var htmlWeb = new HtmlWeb
            {
                OverrideEncoding = Encoding.UTF8
            };


            var htmlDoc = new HtmlDocument();

            //GodLikeHTML.Load(GodLikeClient.OpenRead("http://www.alfa.lt"), Encoding.UTF8);

            //HttpDownloader downloader = new HttpDownloader("http://www.alfa.lt", null, null);
            //GodLikeHTML.LoadHtml(downloader.GetPage());

            //for (int i = 0; i < link.Length; i++)
            //{
            //    Console.WriteLine(link.Length); 
            //    Console.WriteLine(link[i]);
            //}

            DateTime startDate = new DateTime(); //--Time work inicializing

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
            }

            MainEVHF.CreateDir(@"C:\New folder");
            for (int k = 0; k < link.Length; k++)
            {
                WebClient wc = new WebClient
                {
                    Encoding = Encoding.UTF8
                };
                var result = wc.DownloadString(link[k]);
                //Console.WriteLine(result);
                // "printServlet?tkn=MTcxMjU5MDMwMjIxNDMwNzA5ODQsMUhSUkIxWiwxLDE1Mjc1NzcwNDkxMDIsMDA3NDc1MTE="
                // Example #2: Write one string to a text file.
                //string text = "A class is the most powerful data type in C#. Like a structure, " +
                //               "a class defines the data and behavior of the data type. ";
                // WriteAllText creates a file, writes the specified string to the file,
                // and then closes the file.    You do NOT need to call Flush() or Close().
                System.IO.File.WriteAllText($@"C:\New folder\text{k}.html", result);
            }

            for (int m = 0; m < link.Length; m++)
            {
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
                    //htmlDoc.Load($@"C:\New folder\text{m}.html");
                    htmlDoc = htmlWeb.Load($@"C:\New folder\text{m}.html");
                }
                catch (Exception e) 
                {
                    Console.WriteLine("The file could not be read:");
                    Console.WriteLine(e.Message);
                }
                //startDate = DateTime.Now; //--Time work start
                //EVHFList.AddRange(StringToListEVHF(RZEncoding.HTMLToUTF8(htmlDoc.ParsedText)));
                EVHFList.AddRange(StringToListEVHF(htmlDoc.ParsedText));
                Console.WriteLine($"File {m} added");
            }
            
            DateTime endDate = DateTime.Now; //--Time work start
            Console.WriteLine(endDate-startDate);
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
                        //    RZEVHF[13]));
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
                dt0.Rows.Add(new object[] { MainEVHF.EVHFIO, i.Voen, i.Ad, i.Tip, i.Veziyyet, i.Tarix, i.Seriya, i.Nomre, i.EsasQeyd, i.ElaveQeyd, i.EDVsiz, i.EDV, i.Hesab1C, i.MVQeyd });
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

        public int CompareTo(object obj)
        {
            var eVHF = obj as EVHF;
            if (int.Parse(this.Nomre) > int.Parse(eVHF.Nomre))
                return 1;
            else if (int.Parse(this.Nomre) < int.Parse(eVHF.Nomre))
                return -1;
            else return 0;
        }

        //object IEnumerator.Current => throw new NotImplementedException();

        //bool IEnumerator.MoveNext()
        //{
        //    throw new NotImplementedException();
        //}

        //void IEnumerator.Reset()
        //{
        //    throw new NotImplementedException();
        //}

        public static bool operator >(EVHF obj1, EVHF obj2)
        {
            return int.Parse(obj1.Nomre) > int.Parse(obj2.Nomre);
        }
        public static bool operator <(EVHF obj1, EVHF obj2)
        {
            return int.Parse(obj1.Nomre) < int.Parse(obj2.Nomre);
        }
        public object this[int index]
        {
            get
            {
                switch (index)
                {
                    case 0:
                        return IO;
                    case 1:
                        return Voen;
                    case 2:
                        return Ad;
                    case 3:
                        return Tip;
                    case 4:
                        return Veziyyet;
                    case 5:
                        return Tarix;
                    case 6:
                        return Seriya;
                    case 7:
                        return Nomre;
                    case 8:
                        return EsasQeyd;
                    case 9:
                        return ElaveQeyd;
                    case 10:
                        return EDVsiz;
                    case 11:
                        return EDV;
                    case 12:
                        return Hesab1C;
                    case 13:
                        return MVQeyd;
                    default:
                        return 99;
                }
            }
            set
            {
                switch (index)
                {
                    case 0:
                        IO = (string)value;
                        break;
                    case 1:
                        Voen = (string)value;
                        break;
                    case 2:
                        Ad = (string)value;
                        break;
                    case 3:
                        Tip = (string)value;
                        break;
                    case 4:
                        Veziyyet = (string)value;
                        break;
                    case 5:
                        Tarix = (string)value;
                        break;
                    case 6:
                        Seriya = (string)value;
                        break;
                    case 7:
                        Nomre = (string)value;
                        break;
                    case 8:
                        EsasQeyd = (string)value;
                        break;
                    case 9:
                        ElaveQeyd = (string)value;
                        break;
                    case 10:
                        EDVsiz = (string)value;
                        break;
                    case 11:
                        EDV = (string)value;
                        break;
                    case 12:
                        Hesab1C = (string)value;
                        break;
                    case 13:
                        MVQeyd = (string)value;
                        break;
                    default:
                        break;
                }
            }
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
            string insertLink;
            bool tokenExsist = false;
            do
            {
                Console.Clear();
                Console.WriteLine("Pleese inser Link");
                insertLink = Console.ReadLine();
                if (CopyToken(insertLink).Length > 0) tokenExsist = true;
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
            Console.WriteLine($"VOEN: {CopyVoen(insertLink)}");

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
            EVHFsLink = insertLink;
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
                            if (/*link[xlen] == '=' || */link[xlen] == '&') break;
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
                                if (/*link[xlen] == '=' || */link[xlen] == '&') break;
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
            string XVoen = "";
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
                            if (xlen >= link.Length) xlen = link.Length - 1;
                            if (link[xlen] == '=' || link[xlen] == '&') break;
                            if (xlen <= link.Length - 1) XVoen += link[xlen]; else break;
                        } while (true);
                        //XToken = XToken.Remove(XToken.Length - 1, 1);
                    }
                }
            }
            if (XVoen.Length == 0)
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
                                if (xlen <= link.Length - 1) XVoen += link[xlen]; else break;
                            } while (true);
                            //XToken = XToken.Remove(XToken.Length - 1, 1);
                        }
                    }
                }
            }
            //Console.WriteLine(XToken);
            return XVoen;
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
                    @"printServlet?tkn=" + CopyToken(link) +
                    @"&w=2" +
                    @"&v=" +
                    @"&fd=" + strDate + @"000000" +
                    @"&td=" + strDate + @"000000" +
                    @"&s=" +
                    @"&n=" +
                    @"&sw=0" +
                    @"&r=1" +
                    @"&sv=" + EVHFsVOEN;
                //Console.WriteLine(linkArray[i]);
                //linkArray[i] = $"C:\\text{i}.html";
            }
            return linkArray;
        }
        public static void CreateDir(string path)
        {
            // Specify the directory you want to manipulate.
            //path = @"C:\EVHF files";

            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    Console.WriteLine("That path exists already.");
                    return;
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));

                // Delete the directory.
                //di.Delete();
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            finally { }
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