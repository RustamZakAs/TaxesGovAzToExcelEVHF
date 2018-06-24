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

namespace TaxesGovAzToExcel
{
    class MainTaxes
    {
        public static short DocType { get; set; }
        //*****************************************
        public static string TaxesIO { get; set; }
        //*****************************************
        public static int TaxesVeziyyet { get; set; }
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
        public static string EVHFsVOEN { get; set; }
        //*****************************************
        public static string TaxesIlkTarix { get; set; }
        //*****************************************
        public static string TaxesSonTarix { get; set; }
        //*****************************************
        public static void MainMenyu ()
        {
            Console.Clear();

            EVHFsVOEN = "1501069851";

            Console.Write("\nSened növünü seçin: ");
            Console.BackgroundColor = ConsoleColor.Blue;
            DocType = ChangeDocType(Console.CursorLeft, Console.CursorTop); //0 - EVHF   1 - E-Qaimə   2 - Depozit hesabından çıxrış
            Console.ResetColor();

            Console.WriteLine();

            Console.Write("\nHereket növünü seçin: ");
            Console.BackgroundColor = ConsoleColor.Blue;
            TaxesIO = ChangeEVHFIO(Console.CursorLeft, Console.CursorTop);
            Console.ResetColor();

            Console.WriteLine();

            Console.Write("\nSenedler: ");
            Console.BackgroundColor = ConsoleColor.Blue;
            TaxesVeziyyet = ChangeVeziyyet(Console.CursorLeft, Console.CursorTop);
            Console.ResetColor();

            Console.WriteLine();

            string insertLink;
            bool tokenExsist = false;
            do
            {
                //Console.Clear();
                Console.WriteLine("\nSaytin linkini daxil edin:");
                Console.BackgroundColor = ConsoleColor.Blue;
                insertLink = Console.ReadLine();
                Console.ResetColor();
                if (CopyToken(insertLink).Length > 0) tokenExsist = true;

                if (tokenExsist == true)
                {
                    if (insertLink.IndexOf("vroom") > -1)
                    {
                        if (DocType == 0 && insertLink.IndexOf("vroom") > -1)
                        {

                        }
                        else
                        {
                            Console.BackgroundColor = ConsoleColor.Red;
                            Console.WriteLine("Link sehv daxil edilib!");
                            Console.ResetColor();
                            Console.ReadKey();
                            MainMenyu();
                        }
                    }
                    else if (insertLink.IndexOf("eqaime") > -1)
                    {
                        if (DocType == 1 && insertLink.IndexOf("eqaime") > -1)
                        {

                        }
                        else
                        {
                            Console.BackgroundColor = ConsoleColor.Red;
                            Console.WriteLine("Link sehv daxil edilib!");
                            Console.ResetColor();
                            Console.ReadKey();
                            MainMenyu();
                        }
                    }
                    else
                    {
                        Console.BackgroundColor = ConsoleColor.Red;
                        Console.WriteLine("Link sehv daxil edilib!");
                        Console.ResetColor();
                        Console.ReadKey();
                        MainMenyu();
                    }
                }
            } while (!tokenExsist);

            bool errorDetector = false;
            do
            {
                if (errorDetector)
                {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine(" Tarix sehv daxil edilib:");
                    Console.ResetColor();
                }
                Console.Write("\nIlk tarixi daxil edin: YYYYMMDD  ");
                Console.BackgroundColor = ConsoleColor.Blue;
                TaxesIlkTarix = Console.ReadLine();
                Console.ResetColor();
                if (!ChackDate(TaxesIlkTarix)) errorDetector = true;
            } while (!ChackDate(TaxesIlkTarix));
            errorDetector = false;
            do
            {
                if (errorDetector)
                {
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine(" Tarix sehv daxil edilib:");
                    Console.ResetColor();
                }
                Console.Write("\nSon tarixi daxil edin: YYYYMMDD  ");
                Console.BackgroundColor = ConsoleColor.Blue;
                TaxesSonTarix = Console.ReadLine();
                Console.ResetColor();
                if (!ChackDate(TaxesIlkTarix)) errorDetector = true;
            } while (!ChackDate(TaxesSonTarix));

            string voen;
            if (CopyVoen(insertLink).Length == 0) voen = EVHFsVOEN; else voen = CopyVoen(insertLink);
            Console.WriteLine($"\nVOEN: {voen}");

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
        private static string ChangeEVHFIO(int left, int top)
        {
            ConsoleKeyInfo cki;
            int m_ind = 0;
            int m_count = 2;
            var m_list = new string[m_count];
            m_list[0] = "Gelenler      ";
            m_list[1] = "Gönderdiklerim";
            //Console.SetCursorPosition(left, top);
            //Console.WriteLine(m_list[0]);
            do
            {
                Console.SetCursorPosition(left, top);
                Console.Write(m_list[m_ind]);

                cki = Console.ReadKey();
                if (cki.Key == ConsoleKey.DownArrow)
                {
                    m_ind += 1;
                    if (m_ind >= m_count)
                    {
                        m_ind = m_count - 1;
                    }
                }
                if (cki.Key == ConsoleKey.UpArrow)
                {
                    m_ind -= 1;
                    if (m_ind <= 0)
                    {
                        m_ind = 0;
                    }
                }
                if (cki.Key == ConsoleKey.Enter)
                {
                    switch (m_ind)
                    {
                        case 0:
                            return "I";
                        case 1:
                            return "O";
                    }
                }
            } while (true);
        }
        private static short ChangeDocType(int left, int top)
        {
            ConsoleKeyInfo cki;
            int m_ind = 0;
            int m_count = 2;
            var m_list = new string[m_count];
            m_list[0] = "Elektron Vergi Hesab Fakturalar";
            m_list[1] = "Elektron Qaimeler              ";
            //Console.SetCursorPosition(left, top);
            //Console.WriteLine(m_list[0]);
            do
            {
                Console.SetCursorPosition(left, top);
                Console.Write(m_list[m_ind]);

                cki = Console.ReadKey();
                if (cki.Key == ConsoleKey.DownArrow)
                {
                    m_ind += 1;
                    if (m_ind >= m_count)
                    {
                        m_ind = m_count - 1;
                    }
                }
                if (cki.Key == ConsoleKey.UpArrow)
                {
                    m_ind -= 1;
                    if (m_ind <= 0)
                    {
                        m_ind = 0;
                    }
                }
                if (cki.Key == ConsoleKey.Enter)
                {
                    switch (m_ind)
                    {
                        case 0:
                            return 0; //EVHF
                        case 1:
                            return 1; //E-Qaimə
                    }
                }
            } while (true);
        }
        private static int ChangeVeziyyet(int left, int top)
        {
            ConsoleKeyInfo cki;
            int m_ind = 0;
            int m_count = 0;
            string[] m_list = null;
            if (DocType == 0)
            {
                m_count = 4;
                m_list = new string[m_count];
                m_list[0] = "Hamısı          ";
                m_list[1] = "Normal          ";
                m_list[2] = "Ləğv olunmuşlar ";
                m_list[3] = "Dəqiqləşmiş     ";
            }
            else if (DocType == 1)
            {
                m_count = 12;
                m_list = new string[m_count];
                m_list[0]  = "Umumi                                                     ";
                m_list[1]  = "Dəqiqləşmiş                                               ";
                m_list[2]  = "Ləğv edilib              (Qaimə ləğv edilib)              ";
                m_list[3]  = "                         (Təsdiq gözləyən)                ";
                m_list[4]  = "Normal                   (Təsdiqlənmiş)                   ";
                m_list[5]  = "EVHF hazırlanıb          (Faktura hazırlanıb)             ";
                m_list[6]  = "Rədd olunub              (Rədd olunub)                    ";
                m_list[7]  = "EVHF göndərilib          (Faktura göndərilib)             ";
                m_list[8]  = "EVHF ləğv olunub         (Faktura ləğv olunub)            ";
                m_list[9]  = "                         (Sistem tərəfindən təsdiqlənmiş) ";
                m_list[10] = "Sistem EVHF hazırlayıb   (Sistem fakturanı hazırlayıb)    ";
                m_list[11] = "Sistem qaiməni ləğv edib (Sistem qaiməni ləğv edib)       ";
            }
            //Console.SetCursorPosition(left, top);
            //Console.WriteLine(m_list[0]);
            do
            {
                Console.SetCursorPosition(left, top);
                Console.Write(m_list[m_ind]);

                cki = Console.ReadKey();
                if (cki.Key == ConsoleKey.DownArrow)
                {
                    m_ind += 1;
                    if (m_ind >= m_count)
                    {
                        m_ind = m_count - 1;
                    }
                }
                if (cki.Key == ConsoleKey.UpArrow)
                {
                    m_ind -= 1;
                    if (m_ind <= 0)
                    {
                        m_ind = 0;
                    }
                }
                if (cki.Key == ConsoleKey.Enter)
                {
                    return m_ind;
                }
            } while (true);
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
                        for (int t = link.Length - 10; t < link.Length; t++)
                        {
                            XVoen += link[t];
                        }
                        //int x = 0, xlen = 0;
                        //do
                        //{
                        //    xlen = i + Xtemp.Length + x++;
                        //    if (xlen >= link.Length) xlen = link.Length - 1;
                        //    if (link[xlen] == '=' || link[xlen] == '&') break;
                        //    if (xlen <= link.Length - 1) XVoen += link[xlen]; else break;
                        //} while (true);
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
                            for (int t = link.Length - 10; t < link.Length; t++)
                            {
                                XVoen += link[t];
                            }
                            //int x = 0, xlen = 0;
                            //do
                            //{
                            //    xlen = i + Xtemp.Length + x++;
                            //    if (link[xlen] == '=' || link[xlen] == '&') break;
                            //    if (xlen <= link.Length - 1) XVoen += link[xlen]; else break;
                            //} while (true);
                        }
                    }
                }
            }
            return XVoen;
        }
        private static bool ChackDate(string str)
        {
            if (str.Length == 8)
            {
                string str_year = "";
                str_year += str[0];
                str_year += str[1];
                str_year += str[2];
                str_year += str[3];
                int year = int.Parse(str_year);
                if (year <= 2000 || year > (DateTime.Now).Year) return false;

                string str_month = "";
                str_month += str[4];
                str_month += str[5];
                int month = int.Parse(str_month);
                if (month < 1 || month > 12) return false;

                string str_day = "";
                str_day += str[6];
                str_day += str[7];
                int day = int.Parse(str_day);
                if (day < 1 || day > DateTime.DaysInMonth(year, month)) return false;
            }
            else return false;
            return true;
        }
        private static DateTime SQLStrToDate(string str)
        {
            string xyear = "", xmonth = "", xday = "";
            xyear  += str[0];
            xyear  += str[1];
            xyear  += str[2];
            xyear  += str[3];
            xmonth += str[4];
            xmonth += str[5];
            xday   += str[6];
            xday   += str[7];
            DateTime date = new DateTime(int.Parse(xyear), int.Parse(xmonth),int.Parse(xday));
            return date;
        }
        private static string[] CreateLinkArray (string link, string beginDate, string endDate)
        {
            //EVHF    //https://vroom.e-taxes.gov.az/index/index/printServlet?tkn=OTAyMzEyNjI4OTA2OTQ1MTQ1LG51bGwsNCwxNTI4ODI5MjMyNTY2LDAwMTM4OTcx&w=2&v=&fd=20180612000000&td=20180612000000&s=&n=&sw=0&r=1&sv=1501069851
            //E-Qaimə //http://eqaime.e-taxes.gov.az/index/index/printServlet?tkn=OTAyMzEyNjI4OTA2OTQ1MTQ1LG51bGwsMywxNTI4ODI5MDQ3NzgzLDAwMTM4OTcx&w=2&v=&fd=20180612000000&td=20180612000000&s=&n=&sw=0&r=1&sv=1501069851

            DateTime beginDateTime = SQLStrToDate(beginDate);

            TimeSpan difference = SQLStrToDate(endDate).Date - beginDateTime.Date;
            int days = (int)difference.TotalDays + 1;
            Console.WriteLine($"{days} Days");

            string[] linkArray = new string[days];
            
            DateTime tempDateTime;

            string[] sayt = new string[] { "s://vroom", "://eqaime" };

            for (int i = 0; i < days; i++)
            {
                string strDate = "";
                tempDateTime = beginDateTime.AddDays(i);
                strDate = strDate + tempDateTime.Year.ToString();
                strDate += tempDateTime.Month.ToString().Length == 1 ? $"0{tempDateTime.Month.ToString()}" : $"{tempDateTime.Month.ToString()}";
                strDate += tempDateTime.Day.ToString().Length == 1 ? $"0{tempDateTime.Day.ToString()}" : $"{tempDateTime.Day.ToString()}";

                linkArray[i] = @"http"+ sayt[DocType] + ".e-taxes.gov.az/index/index/" +
                    @"printServlet?tkn=" + CopyToken(link) +
                    @"&w=2" +
                    @"&v="  +
                    @"&fd=" + strDate + @"000000" +
                    @"&td=" + strDate + @"000000" +
                    @"&s="  +
                    @"&n="  +
                    @"&sw=" + (TaxesIO == "I" ? "0" : "1") +
                    @"&r=1" +
                    @"&sv=" + EVHFsVOEN;
            }

            int xleft = Console.CursorLeft, xtop = Console.CursorTop;
            for (int i = 0; i < linkArray.Length; i++)
            {
                try
                {
                    Console.SetCursorPosition(xleft, xtop);
                    Console.WriteLine($"{i + 1} link oxunur");
                    if (CheckLink(linkArray[i])) continue;
                    else throw new Exception("Линк не отвечает");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error link: {e}");
                }
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
        public static bool CheckLink(string link)
        {
            //var htmlWeb = new HtmlWeb
            //{
            //    OverrideEncoding = Encoding.UTF8
            //};
            //var htmlDoc = new HtmlDocument();
            Stream stream = new MemoryStream();
            using (StreamWriter sw = new StreamWriter(stream))
            {
                sw.Write(link);                
            }
            WebClient wc = new WebClient
            {
                Encoding = Encoding.UTF8
            };
            string result;
            try
            {
                result = wc.DownloadString(link);
            }
            catch (Exception)
            {
                return false;
                throw;
            }
            if (result.Length > 0) return true;
            return false;
        }

        public static void Main(string[] args)
        {
            MainMenyu();
            string[] linrArray = CreateLinkArray(EVHFsLink, TaxesIlkTarix, TaxesSonTarix);

            List<EVHF> EVHFlist = new List<EVHF>();
            List<EQaime> EQlist = new List<EQaime>();
            if (DocType == 0)
            {
                EVHF.RZLoadFromTaxes(ref EVHFlist, linrArray);
                EVHF.CreateExcel(ref EVHFlist);
            }
            else
            {
                EQaime.RZLoadFromTaxes(ref EQlist, linrArray);
                EQaime.CreateExcel(ref EQlist);
            }
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