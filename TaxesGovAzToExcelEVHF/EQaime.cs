using ExportToExcel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace TaxesGovAzToExcelEVHF
{
    public class EQaime : EVHF
    {
        public string EDVcelb { get; set; }
        public string EDVcelbNo { get; set; }
        public string EDVcelb0 { get; set; }

        public EQaime() { }

        public EQaime(EQaime eQaime) {
            IO = eQaime.IO;
            Voen = eQaime.Voen;
            Ad = eQaime.Ad;
            Tip = eQaime.Tip;
            Veziyyet = eQaime.Veziyyet;
            Tarix = eQaime.Tarix;
            Seriya = eQaime.Seriya;
            Nomre = eQaime.Nomre;
            EsasQeyd = eQaime.EsasQeyd;
            ElaveQeyd = eQaime.ElaveQeyd;
            EDVsiz = eQaime.EDVsiz;
            EDV = eQaime.EDV;
            EDVcelb = eQaime.EDVcelb;
            EDVcelbNo = eQaime.EDVcelbNo;
            EDVcelb0 = eQaime.EDVcelb0;
            Hesab1C = eQaime.Hesab1C;
            MVQeyd = eQaime.MVQeyd;
        }

        public EQaime(string iO, string voen, string ad, string tip, string veziyyet, string tarix, string seriya, string nomre, string esasQeyd, string elaveQeyd, string eDVsiz, string eDV, string eDVcelb, string eDVcelbNo, string eDVcelb0, string hesab1C, string mVQeyd)
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
            EDVcelb = eDVcelb;
            EDVcelbNo = eDVcelbNo;
            EDVcelb0 = eDVcelb0;
            Hesab1C = hesab1C;
            MVQeyd = mVQeyd;
        }

        public static void RZLoadFromTaxes(ref List<EQaime> EQlist, string[] link)
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

            var htmlWeb = new HtmlWeb
            {
                OverrideEncoding = Encoding.UTF8
            };
            var htmlDoc = new HtmlDocument();

            //DateTime startDate = new DateTime(); //--Time work inicializing

            var temp = Path.GetTempFileName();
            var tempFile = temp.Replace(Path.GetExtension(temp), ".html");

            MainTaxes.CreateDir(@"C:\RZUploadingTaxesDocuments");

            for (int i = 0; i < link.Length; i++)
            {
                try
                {
                    //Запись в Stream всех ссылок для последующего отображения
                    using (StreamWriter sw = new StreamWriter(tempFile))
                    {
                        for (int j = 0; j < link.Length; j++)
                        {
                            sw.Write($"{link[j]}<br>");
                        }
                    }
                    WebClient wc = new WebClient
                    {
                        Encoding = Encoding.UTF8
                    };
                    var result = wc.DownloadString(link[i]);

                    string type;
                    if (MainTaxes.DocType == 0) type = "EVHF("; else type = "E-Qaime(";
                    System.IO.File.WriteAllText($@"C:\RZUploadingTaxesDocuments\{type}{i}).html", result);
                    Console.WriteLine($"File {i} created");
                }
                catch (Exception e)
                {
                    Console.WriteLine("---Qoshulamadi---");
                    Console.WriteLine(e);
                    throw;
                }
            }
            // From File
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
                    string type;
                    if (MainTaxes.DocType == 0) type = "EVHF("; else type = "E-Qaime(";
                    htmlDoc = htmlWeb.Load($@"C:\RZUploadingTaxesDocuments\{type}{m}).html");
                }
                catch (Exception e)
                {
                    Console.WriteLine("---Fayl oxunmuyor---");
                    Console.WriteLine(e.Message);
                }
                //startDate = DateTime.Now; //--Time work start
                //EVHFList.AddRange(StringToListEVHF(RZEncoding.HTMLToUTF8(htmlDoc.ParsedText)));
                EQlist.AddRange(StringToListEQ(htmlDoc.ParsedText));
                Console.WriteLine($"File {m} added");
            }

            //DateTime endDate = DateTime.Now; //--Time work start
            //Console.WriteLine(endDate - startDate);  // raznica vo vremeni raboti
            Process.Start(new ProcessStartInfo(tempFile));
        }
        public static List<EQaime> StringToListEQ(string str)
        {
            var RZEQList = new List<EQaime>();
            var RZEQ = new EQaime();

            int j = 0, k = 0, count = 0;
            for (int i = 0; i < str.Length; i++)
            {
                string tempDocx = "";
                for (; j < MainTaxes.TextForBegin.Length; j++)
                {
                    tempDocx += str[(i + j) >= str.Length - 1 ?
                        str.Length - 1 : (i + j)];
                }
                if (tempDocx == MainTaxes.TextForBegin)
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
                        RZEQ/*[0]*/.IO = MainTaxes.TaxesIO;
                        RZEQ/*[1]*/.Voen = Xvalue;
                    }
                    if (count == 2) RZEQ/*[2]*/.Ad = Xvalue;
                    if (count == 3) RZEQ/*[3]*/.Tip = Xvalue;
                    if (count == 4) RZEQ/*[4]*/.Veziyyet = Xvalue;
                    if (count == 5) RZEQ/*[5]*/.Tarix = Xvalue;
                    if (count == 6) RZEQ/*[6]*/.Seriya = Xvalue;
                    if (count == 7) RZEQ/*[7]*/.Nomre = Xvalue;
                    if (count == 8) RZEQ/*[8]*/.EsasQeyd = Xvalue;
                    if (count == 9) RZEQ/*[9]*/.ElaveQeyd = Xvalue;
                    if (count == 10)
                    {
                        //Xvalue = Xvalue.Replace(".", ",");
                        //RZEVHF.EDVsiz = decimal.Parse(Xvalue);
                        RZEQ/*[10]*/.EDVsiz = Xvalue;
                    }
                    if (count == 11)
                    {
                        //Xvalue = Xvalue.Replace(".", ",");
                        //RZEVHF.EDV = decimal.Parse(Xvalue);
                        RZEQ/*[11]*/.EDV = Xvalue;
                    }
                    if (count == 12)
                    {
                        RZEQ.EDVcelb = Xvalue;
                    }
                    if (count == 13)
                    {
                        RZEQ.EDVcelbNo = Xvalue;
                    }
                    if (count == 14)
                    {
                        RZEQ.EDVcelb0 = Xvalue;
                        RZEQ/*[12]*/.Hesab1C = (MainTaxes.TaxesIO == "I" ? "531.1" : "211.3");
                        RZEQ/*[13]*/.MVQeyd = "";
                        //Console.WriteLine(RZEVHF.ToString());
                        RZEQList.Add(new EQaime(RZEQ));
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
            return RZEQList;
        }
        public static void CreateExcel(ref List<EQaime> EQs)
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
            DataSet ds = ExportToExcel(ref EQs);
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
        public static DataSet ExportToExcel(ref List<EQaime> EQs)
        {
            //  Create a sample DataSet, containing three DataTables.
            //  (Later, this will save to Excel as three Excel worksheets.)
            DataSet ds0 = new DataSet();
            //  Create the first table of sample data
            DataTable dt0 = new DataTable("E-Qaimə");
            //dt0.Rows.Add(new object[] { "E-Qaimə siyahısı" });
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
            dt0.Columns.Add("ƏDV-yə cəlb edilən", Type.GetType("System.String"));
            dt0.Columns.Add("ƏDV-yə cəlb edilməyən", Type.GetType("System.String"));
            dt0.Columns.Add("ƏDV-yə “0” dərəcə ilə cəlb edilən", Type.GetType("System.String"));
            dt0.Columns.Add("1C", Type.GetType("System.String"));
            dt0.Columns.Add("Malverən qeyd", Type.GetType("System.String"));

            foreach (var i in EQs)
            {
                dt0.Rows.Add(new object[] { MainTaxes.TaxesIO,
                    i.Voen,
                    i.Ad,
                    i.Tip,
                    i.Veziyyet,
                    i.Tarix,
                    i.Seriya,
                    i.Nomre,
                    i.EsasQeyd,
                    i.ElaveQeyd,
                    i.EDVsiz,
                    i.EDV,
                    i.EDVcelb,
                    i.EDVcelbNo,
                    i.EDVcelb0,
                    i.Hesab1C,
                    i.MVQeyd }
                );
            }
            ds0.Tables.Add(dt0);
            return ds0;
        }
    }
}
