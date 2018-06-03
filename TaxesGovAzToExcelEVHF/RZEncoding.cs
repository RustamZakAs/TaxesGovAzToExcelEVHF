using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TaxesGovAzToExcelEVHF
{
    class RZEncoding
    {
        static public string HTMLToUTF8 (string str)
        {
            //---AZ---
            str = str.Replace("ЖЏ", "Ə");
            str = str.Replace("Й™", "ə");
            str = str.Replace("Г–", "Ö");
            str = str.Replace("Г¶", "ö");
            str = str.Replace("Дћ", "Ğ");
            str = str.Replace("Дџ", "ğ");
            str = str.Replace("Д°", "İ");
            str = str.Replace("Д±", "ı");
            str = str.Replace("Гњ", "Ü");
            str = str.Replace("Гј", "ü");
            str = str.Replace("Г‡", "Ç");
            str = str.Replace("Г§", "ç");
            str = str.Replace("Ећ", "Ş");
            str = str.Replace("Еџ", "ş");
            //---SIMBOL---
            str = str.Replace("вЂњ", "“");
            str = str.Replace("вЂќ", "”");
            str = str.Replace("&", "&");
            //---RU---
            str = str.Replace("Р‘", "Б");
            str = str.Replace("Р“", "Г");
            str = str.Replace("Рљ", "К");
            str = str.Replace("Рњ", "М");
            str = str.Replace("Рћ", "О");
            str = str.Replace("РЎ", "С");
            str = str.Replace("Р°", "а");
            str = str.Replace("РІ", "в");
            str = str.Replace("Рі", "г");
            str = str.Replace("Рґ", "д");
            str = str.Replace("Рµ", "е");
            str = str.Replace("Р·", "з");
            str = str.Replace("Рё", "и");
            str = str.Replace("Р№", "й");
            str = str.Replace("Р»", "л");
            str = str.Replace("Рј", "м");
            str = str.Replace("РЅ", "н");
            str = str.Replace("Рѕ", "о");
            str = str.Replace("СЂ", "р");
            str = str.Replace("СЃ", "с");
            str = str.Replace("Сѓ", "у");
            str = str.Replace("С€", "ш");
            str = str.Replace("СЏ", "я");

            return str;
        }
    }
}
