using BenchmarkDotNet.Running;
using System.Diagnostics;

namespace TextReplaceAPI.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Naming convention for benchmark classes:
            // <Method>_<FilesBeingBenchmarked>_<ExtraParameters>

            BenchmarkRunner.Run<Benchmarks>();

            // PerformBenchmarksDry();
        }

        static void PerformBenchmarksDry()
        {
            var bm = new Benchmarks();
            bm.Setup();

            Debug.WriteLine("Number of replacements made:");
            Debug.WriteLine("Bible: " + bm.GetBibleReplacements());
            Debug.WriteLine("LOTR: " + bm.GetLotrReplacements());
            Debug.WriteLine("Harry Potter: " + bm.GetHarryPotterReplacements());
            Debug.WriteLine("Resumes: " + bm.GetResumesReplacements());
            Debug.WriteLine("Spreadsheet: " + bm.GetSpreadsheetReplacements());
        }
    }
}
