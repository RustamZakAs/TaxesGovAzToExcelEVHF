using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TaxesGovAzToExcelEVHF
{
    public class EQaime : EVHF
    {
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

        public string EDVcelb { get; set; }
        public string EDVcelbNo { get; set; }
        public string EDVcelb0 { get; set; }


    }
}
