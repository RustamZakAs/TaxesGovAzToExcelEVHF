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

namespace TaxesGovAzToExcelEVHF
{
    class MainEVHF
    {
        //*****************************************
        public static string EVHFIO { get; set; }
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
            EVHFIO = "O";
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
            //        @"printServlet?tkn=" + CopyToken(link) +
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
                            if (link[xlen] == '&') break;
                            if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                        } while (true);
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
                                if (link[xlen] == '&') break;
                                if (xlen <= link.Length - 1) XToken += link[xlen]; else break;
                            } while (true);
                        }
                    }
                }
            }
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
                        }
                    }
                }
            }
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
                    @"&sw=" + (EVHFIO == "I" ? "0" : "1") +
                    @"&r=1" +
                    @"&sv=" + EVHFsVOEN;
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