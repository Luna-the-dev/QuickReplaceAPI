using BenchmarkDotNet.Attributes;

namespace TextReplaceAPI.Benchmarks
{
    [MemoryDiagnoser]
    public class Benchmarks
    {
        // relative path when running in debug: ../../../
        // relative path when running in release: ../../../../../../../

        public const string REPLACEMENTS = "../../../../../../../Common/replacements.csv";

        private const string BIBLE_SRC = "../../../../../../../Common/bible.docx";
        private const string LOTR_SRC = "../../../../../../../Common/lord-of-the-rings.docx";
        private const string HARRYPOTTER_SRC = "../../../../../../../Common/harrypotter.docx";
        private const string RESUME_1_SRC = "../../../../../../../Common/resume-1.docx";
        private const string RESUME_2_SRC = "../../../../../../../Common/resume-2.docx";
        private const string SPREADSHEET_SRC = "../../../../../../../Common/financial-sample.xlsx";

        private const string BIBLE_OUT = "../../../../../../../Outputs/bible.docx";
        private const string LOTR_OUT = "../../../../../../../Outputs/lord-of-the-rings.docx";
        private const string HARRYPOTTER_OUT = "../../../../../../../Outputs/harrypotter.docx";
        private const string RESUME_1_OUT = "../../../../../../../Outputs/Resumes/resume-1.docx";
        private const string RESUME_2_OUT = "../../../../../../../Outputs/Resumes/resume-2.docx";
        private const string SPREADSHEET_OUT = "../../../../../../../Outputs/financial-sample.xlsx";

        private List<string> resumeSources = new List<string>();
        private List<string> resumeOutputs = new List<string>();

        private QuickReplace? qrBible;
        private QuickReplace? qrLotr;
        private QuickReplace? qrHarryPotter;
        private QuickReplace? qrHundredResumes;
        private QuickReplace? qrSpreadsheet;

        private QuickReplace? qrMatcherBible;
        private QuickReplace? qrMatcherLotr;
        private QuickReplace? qrMatcherHarryPotter;
        private QuickReplace? qrMatcherHundredResumes;
        private QuickReplace? qrMatcherSpreadsheet;

        [GlobalSetup]
        public void Setup()
        {
            // build the lists of resumes
            string resume1Directory = Path.GetDirectoryName(RESUME_1_OUT) ?? "";
            string resume1FileName = Path.GetFileNameWithoutExtension(RESUME_1_OUT);
            string resume1Extension = Path.GetExtension(RESUME_1_OUT);

            string resume2Directory = Path.GetDirectoryName(RESUME_2_OUT) ?? "";
            string resume2FileName = Path.GetFileNameWithoutExtension(RESUME_2_OUT);
            string resume2Extension = Path.GetExtension(RESUME_2_OUT);

            for (int i = 0; i < 50; i++)
            {
                resumeSources.Add(RESUME_1_SRC);
                resumeSources.Add(RESUME_2_SRC);

                resumeOutputs.Add($"{resume1Directory}/{resume1FileName}_{i+1}{resume1Extension}");
                resumeOutputs.Add($"{resume1Directory}/{resume2FileName}_{i+1}{resume2Extension}");
            }

            // build the QuickReplace objects without pregenerated matchers
            qrBible = new QuickReplace(REPLACEMENTS, [BIBLE_SRC], [BIBLE_OUT]);
            qrLotr = new QuickReplace(REPLACEMENTS, [LOTR_SRC], [LOTR_OUT]);
            qrHarryPotter = new QuickReplace(REPLACEMENTS, [HARRYPOTTER_SRC], [HARRYPOTTER_OUT]);
            qrHundredResumes = new QuickReplace(REPLACEMENTS, resumeSources, resumeOutputs);
            qrSpreadsheet = new QuickReplace(REPLACEMENTS, [SPREADSHEET_SRC], [SPREADSHEET_OUT]);

            // build the QuickReplace objects with pregenerated matchers
            qrMatcherBible = new QuickReplace(REPLACEMENTS, [BIBLE_SRC], [BIBLE_OUT], preGenerateMatcher: true, caseSensitive: false);
            qrMatcherLotr = new QuickReplace(REPLACEMENTS, [LOTR_SRC], [LOTR_OUT], preGenerateMatcher: true, caseSensitive: false);
            qrMatcherHarryPotter = new QuickReplace(REPLACEMENTS, [HARRYPOTTER_SRC], [HARRYPOTTER_OUT], preGenerateMatcher: true, caseSensitive: false);
            qrMatcherHundredResumes = new QuickReplace(REPLACEMENTS, resumeSources, resumeOutputs, preGenerateMatcher: true, caseSensitive: false);
            qrMatcherSpreadsheet = new QuickReplace(REPLACEMENTS, [SPREADSHEET_SRC], [SPREADSHEET_OUT], preGenerateMatcher: true, caseSensitive: false);
        }

        // Naming convention for benchmark classes:
        // <Method>_<FilesBeingBenchmarked>_<ExtraParameters>

        //////////////////////
        // NO PREGENERATION //
        //////////////////////

        [Benchmark]
        public void PerformReplacements_Bible_NoPreGeneration()
        {
            if (qrBible == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrBible.PerformReplacements(false, false, false);
            File.Delete(qrBible.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_Lotr_NoPreGeneration()
        {
            if (qrLotr == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrLotr.PerformReplacements(false, false, false);
            File.Delete(qrLotr.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_HarryPotter_NoPreGeneration()
        {
            if (qrHarryPotter == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrHarryPotter.PerformReplacements(false, false, false);
            File.Delete(qrHarryPotter.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_HundredResumes_NoPreGeneration()
        {
            if (qrHundredResumes == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrHundredResumes.PerformReplacements(false, false, false);
            foreach (var sourceFile in qrHundredResumes.SourceFiles)
            {
                File.Delete(sourceFile.OutputFileName);
            }
        }

        [Benchmark]
        public void PerformReplacements_Spreadsheet_NoPreGeneration()
        {
            if (qrSpreadsheet == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrSpreadsheet.PerformReplacements(false, false, false);
            File.Delete(qrSpreadsheet.SourceFiles[0].OutputFileName);
        }

        ///////////////////
        // PREGENERATION //
        ///////////////////

        [Benchmark]
        public void PerformReplacements_Bible_PreGeneration()
        {
            if (qrMatcherBible == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrMatcherBible.PerformReplacements(false, false, false);
            File.Delete(qrMatcherBible.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_Lotr_PreGeneration()
        {
            if (qrMatcherLotr == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrMatcherLotr.PerformReplacements(false, false, false);
            File.Delete(qrMatcherLotr.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_HarryPotter_PreGeneration()
        {
            if (qrMatcherHarryPotter == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrMatcherHarryPotter.PerformReplacements(false, false, false);
            File.Delete(qrMatcherHarryPotter.SourceFiles[0].OutputFileName);
        }

        [Benchmark]
        public void PerformReplacements_HundredResumes_PreGeneration()
        {
            if (qrMatcherHundredResumes == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrMatcherHundredResumes.PerformReplacements(false, false, false);
            foreach (var sourceFile in qrMatcherHundredResumes.SourceFiles)
            {
                File.Delete(sourceFile.OutputFileName);
            }
        }

        [Benchmark]
        public void PerformReplacements_Spreadsheet_PreGeneration()
        {
            if (qrMatcherSpreadsheet == null)
            {
                throw new NullReferenceException("The QuickReplace object is null.");
            }

            qrMatcherSpreadsheet.PerformReplacements(false, false, false);
            File.Delete(qrMatcherSpreadsheet.SourceFiles[0].OutputFileName);
        }

        /////////////////////
        // UTILITY METHODS //
        /////////////////////

        public int GetBibleReplacements()
        {
            if (qrMatcherBible == null)
            {
                return -1;
            }

            PerformReplacements_Bible_PreGeneration();
            return qrMatcherBible.SourceFiles[0].NumOfReplacements;
        }

        public int GetLotrReplacements()
        {
            if (qrMatcherLotr == null)
            {
                return -1;
            }

            PerformReplacements_Lotr_PreGeneration();
            return qrMatcherLotr.SourceFiles[0].NumOfReplacements;
        }

        public int GetHarryPotterReplacements()
        {
            if (qrMatcherHarryPotter == null)
            {
                return -1;
            }

            PerformReplacements_HarryPotter_PreGeneration();
            return qrMatcherHarryPotter.SourceFiles[0].NumOfReplacements;
        }

        public int GetResumesReplacements()
        {
            if (qrMatcherHundredResumes == null)
            {
                return -1;
            }

            PerformReplacements_HundredResumes_PreGeneration();

            int numOfReplacements = 0;
            foreach (var sourceFile in qrMatcherHundredResumes.SourceFiles)
            {
                numOfReplacements += sourceFile.NumOfReplacements;
            }
            return numOfReplacements;
        }

        public int GetSpreadsheetReplacements()
        {
            if (qrMatcherSpreadsheet == null)
            {
                return -1;
            }

            PerformReplacements_Spreadsheet_PreGeneration();
            return qrMatcherSpreadsheet.SourceFiles[0].NumOfReplacements;
        }
    }
}
