using Kros.KORM;
using Kros.KORM.Query.MsAccess;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace K4Extractor
{
    class Program
    {
        static void Main(string[] args)
        {
            MsAccessQueryProviderFactory.Register();
            var con = @"TopSecret";
            var count = 0;
            using (var korm = new Database(con, "System.Data.OleDb"))
            using (var writer = new StreamWriter("data.csv"))
            using (var testWriter = new StreamWriter("testing.csv"))
            {
                writer.WriteLine("TV;Description");
                testWriter.WriteLine("TV;Description");
                foreach (var item in korm.Query<Polozky>()
                        .Where(p => p.TypVety == "K" || p.TypVety == "M")
                        .OrderBy(p => p.PopisUplny)
                        .AsEnumerable()
                        .Select(p => $"{p.TypVety};{p.PopisUplny.Replace(";", string.Empty)}"))
                {
                    count++;
                    if (count % 1000 == 0)
                    {
                        testWriter.WriteLine(item);
                    }
                    else
                    {
                        writer.WriteLine(item);
                    }
                }

                writer.Flush();
            }
        }

        public class Polozky
        {
            public string TypVety { get; set; }

            public string PopisUplny { get; set; }
        }

    }
}
