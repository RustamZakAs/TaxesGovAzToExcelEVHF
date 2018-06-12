using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TaxesGovAzToExcelEVHF
{
    class RZComboBox
    {
        public string[] Input { get; set; }
        public string[] Output { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Select_id { get; set; }
        public int Select_in { get; set; }
        public int Select_out { get; set; }

        public static string Change(string[] view, string[] output,  int left, int top, int view_id)
        {
            try
            {
                if (view.Length != output.Length)
                {
                    throw new Exception("Массивы не равны!!!");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Ошибка: " + e.Message);

            }
            Console.ReadLine();


            ConsoleKeyInfo cki;
            int m_ind = 0;
            int m_count = 2;
            //var m_list = new string[m_count];
            //m_list[0] = "Gelenler      ";
            //m_list[1] = "Gönderdiklerim";
            //Console.SetCursorPosition(left, top);
            //Console.WriteLine(view[0]);
            do
            {
                Console.SetCursorPosition(left, top);
                Console.WriteLine(view[m_ind]);

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
    }
}
