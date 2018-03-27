using System;
using System.IO;
using System.Linq;
using Kros.KORM;
using Kros.KORM.Query.MsAccess;

namespace K4Extractor
{
    class Program
    {
        static void Main(string[] args)
        {
            MsAccessQueryProviderFactory.Register();
            var con = @"Provider=Microsoft.ACE.OLEDB.12.0;OLE DB Services=-4;Data Source='D:\K4Data\Databazy\ÚRS PRAHA 2018 01.KD';User ID=why;Password=' ';Jet OLEDB:System database=D:\K4Data\System.mdw;Mode=Share Deny None;Locale Identifier=1029;Jet OLEDB:Database Password="";Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Engine Type=5;Persist Security Info=TRUE;";

            using(var korm = new Database(con, "System.Data.OleDb"))
            using(var writer = new StreamWriter("data.csv"))
            {
                foreach (var item in korm.Query<Polozka>()
                        .Where(p => p.TypVety == "K" || p.TypVety == "M")
                        .Select(p => $"{p.TypVety};{p.PopisUplny}"))
                {
                    writer.WriteLine(item);
                }

                writer.Flush();
            }
        }

        public class Polozka
        {
            public string TypVety { get; set; }

            public string PopisUplny { get; set; }
        }
    }
}