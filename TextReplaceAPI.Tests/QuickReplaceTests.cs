using TextReplaceAPI.DataTypes;
using TextReplaceAPI.Exceptions;
using TextReplaceAPI.Tests.Common;

namespace TextReplaceAPI.Tests
{
    public class QuickReplaceTests
    {
        private static readonly string RelativeReplacementsPath = "../../../MockFiles/Replacements/";
        private static readonly string RelativeSourcesPath = "../../../MockFiles/Sources/";
        private static readonly string RelativeMockOutputsPath = "../../../MockFiles/Outputs/";
        private static readonly string RelativeGeneratedFilePath = "../../GeneratedTestFiles/QuickReplaceTests/";

        [Fact]
        public void QuickReplace_ConstructorReplacementsFromDict_ReplacePhrasesAndSourceFilesInitialized()
        {
            // Arrange
            var replacePhrases = new Dictionary<string, string>
            {
                {"basic-text", "basic-text1"},
                {"text with whitespace", "text with whitespace 1"},
                {"text,with,commas", "text,with,commas,1"},
                {"text,with\",commas\"and\"quotes,\"", "text,with\",commas\"and\"quotes,\"1"},
                {"text;with;semicolons", "text;with;semicolons1"}
            };

            var sourceFileNames = new List<string>()
            {
                "source-file-name.txt",
                "source-file-name.csv",
                "source-file-name.tsv",
                "source-file-name.docx",
                "source-file-name.xlsx"
            };

            var outputFileNames = new List<string>()
            {
                "output-file-name.txt",
                "output-file-name.csv",
                "output-file-name.tsv",
                "output-file-name.docx",
                "output-file-name.xlsx"
            };

            // Act
            var QuickReplace = new QuickReplace(replacePhrases, sourceFileNames, outputFileNames);

            // Assert
            Assert.Equal(replacePhrases, QuickReplace.ReplacePhrases);

            Assert.Equal(5, QuickReplace.SourceFiles.Count);
            for (int i = 0; i < sourceFileNames.Count; i++)
            {
                Assert.Equal(sourceFileNames[i], QuickReplace.SourceFiles[i].SourceFileName);
                Assert.Equal(outputFileNames[i], QuickReplace.SourceFiles[i].OutputFileName);
                Assert.Equal(-1, QuickReplace.SourceFiles[i].NumOfReplacements);
            }
        }

        [Theory]
        [InlineData("replacements-abbreviations.csv")]
        [InlineData("replacements-abbreviations.tsv")]
        [InlineData("replacements-abbreviations.xlsx")]
        [InlineData("replacements-abbreviations-pound-delimiter.txt")]
        [InlineData("replacements-abbreviations-semicolon-delimiter.txt")]
        public void QuickReplace_ConstructorReplacementsFromFileName_ReplacePhrasesAndSourceFilesInitialized(string replacementsFileName)
        {
            // Arrange
            replacementsFileName = RelativeReplacementsPath + replacementsFileName;

            var sourceFileNames = new List<string>()
            {
                "source-file-name.txt",
                "source-file-name.csv",
                "source-file-name.tsv",
                "source-file-name.docx",
                "source-file-name.xlsx"
            };

            var outputFileNames = new List<string>()
            {
                "output-file-name.txt",
                "output-file-name.csv",
                "output-file-name.tsv",
                "output-file-name.docx",
                "output-file-name.xlsx"
            };

            var key = "Above-mentioned";
            var expected = "abv-mntnd";

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames);

            // Assert
            Assert.Equal(expected, QuickReplace.ReplacePhrases[key]);

            for (int i = 0; i < sourceFileNames.Count; i++)
            {
                Assert.Equal(sourceFileNames[i], QuickReplace.SourceFiles[i].SourceFileName);
                Assert.Equal(outputFileNames[i], QuickReplace.SourceFiles[i].OutputFileName);
                Assert.Equal(-1, QuickReplace.SourceFiles[i].NumOfReplacements);
            }
        }

        [Fact]
        public void QuickReplace_ConstructorGenerateMatcher_AhoCorasickMatcherInitialized()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";

            var sourceFileNames = new List<string>() { "source-file-name.txt" };
            var outputFileNames = new List<string>() { "output-file-name.txt" };

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames, preGenerateMatcher: true, caseSensitive: true);

            // Assert
            Assert.True(QuickReplace.IsMatcherCreated());
        }

        [Theory]
        [InlineData("invalid-filetype.tmp", "source-resume.txt", "output-resume.txt")]
        [InlineData("replacements-abbreviations.csv", "invalid-filetype.tmp", "source-resume.txt")]
        [InlineData("replacements-abbreviations.csv", "source-resume.txt", "invalid-filetype.tmp")]
        public void QuickReplace_InvalidFileType_ThrowsInvalidFileTypeException(string replacementsFileName, string sourceFileName, string outputFileName)
        {
            // Arrange
            replacementsFileName = RelativeReplacementsPath + replacementsFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { outputFileName };

            // Act and Assert
            Assert.Throws<InvalidFileTypeException>(() => new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames));
        }

        [Fact]
        public void QuickReplace_EmptyReplacements_ThrowsInvalidOperationException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-empty.csv";

            var sourceFileNames = new List<string>() { "source-file-name.txt" };
            var outputFileNames = new List<string>() { "output-file-name.txt" };

            // Act and Assert
            Assert.Throws<InvalidOperationException>(() => new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames));
        }

        [Fact]
        public void QuickReplace_EmptyReplacePhraseItem1_ThrowsInvalidOperationException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-empty-item1.csv";

            var sourceFileNames = new List<string>() { "source-file-name.txt" };
            var outputFileNames = new List<string>() { "output-file-name.txt" };

            // Act and Assert
            Assert.Throws<InvalidOperationException>(() => new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames));
        }

        [Fact]
        public void QuickReplace_DifferentSourceAndOutputFileLengths_ThrowsArgumentException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";

            var sourceFileNames = new List<string>() { "source-file-name.txt", "source-file-name.docx" };

            var outputFileNames = new List<string>() { "output-file-name.txt" };

            // Act and Assert
            Assert.Throws<ArgumentException>(() => new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames));
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public void ReplaceNonStatic_ValidFiles_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";

            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames);
            var actual = QuickReplace.Replace(wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public async Task ReplaceAsyncNonStatic_ValidFiles_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";

            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames);
            var actual = await QuickReplace.ReplaceAsync(wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public void ReplaceNonStatic_ValidFilesNonNullMatcher_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames, preGenerateMatcher: true, caseSensitive: false);
            var matcherActual = QuickReplace.IsMatcherCreated();
            var actual = QuickReplace.Replace(wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(matcherActual);
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public async Task ReplaceAsyncNonStatic_ValidFilesNonNullMatcher_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            // Act
            var QuickReplace = new QuickReplace(replacementsFileName, sourceFileNames, outputFileNames, preGenerateMatcher: true, caseSensitive: false);
            var matcherActual = QuickReplace.IsMatcherCreated();
            var actual = await QuickReplace.ReplaceAsync(wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(matcherActual);
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(QuickReplace.SourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public void ReplaceStatic_ValidFiles_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public async Task ReplaceAsyncStatic_ValidFiles_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = await QuickReplace.ReplaceAsync(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Fact]
        public void ReplaceStatic_EmptySourceFiles_ThrowsArgumentException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            var generatedFileName = RelativeGeneratedFilePath + "output-resume.txt";

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var emptySourceFiles = new List<SourceFile>();

            if (File.Exists(generatedFileName))
            {
                File.Delete(generatedFileName);
            }

            // Act and Assert
            Assert.Throws<ArgumentException>(() =>
                QuickReplace.Replace(replacePhrases, emptySourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false, throwExceptions: false));
            Assert.Throws<ArgumentException>(() =>
                QuickReplace.Replace(replacePhrases, emptySourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false, throwExceptions: true));

            Assert.False(File.Exists(generatedFileName));
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume.txt")]
        [InlineData("source-resume.csv", "output-resume.csv")]
        [InlineData("source-resume.tsv", "output-resume.tsv")]
        [InlineData("source-resume.docx", "output-resume.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample.xlsx")]
        public void ReplaceStatic_OutputFileIsInUse_ThrowsIOException(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            if (File.Exists(generatedFileName) == false)
            {
                File.Create(generatedFileName).Dispose();
            }

            // open a stream writer on the generated file to make sure it cant be written to
            using var writer = new StreamWriter(generatedFileName);

            // Act and Assert
            Assert.Throws<IOException>(() => QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false));
        }

        [Theory]
        [InlineData("source-resume.txt", "output-financial-sample.xlsx")]
        [InlineData("source-financial-sample.xlsx", "output-resume.txt")]
        [InlineData("source-resume.docx", "output-financial-sample.xlsx")]
        [InlineData("source-financial-sample.xlsx", "output-resume.docx")]
        public void ReplaceStatic_InvalidFileTypeConversion_ThrowsNotSupportedException(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Normal/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act and Assert
            Assert.Throws<NotSupportedException>(() =>
                QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false));
        }

        [Theory]
        [InlineData("source-resume.txt", "output-financial-sample.xlsx")]
        [InlineData("output-financial-sample.xlsx", "source-resume.txt")]
        [InlineData("source-resume.docx", "output-financial-sample.xlsx")]
        [InlineData("output-financial-sample.xlsx", "source-resume.docx")]
        public void ReplaceStatic_InvalidFileTypeConversionThrowExceptionsFalse_ReturnsFalseReplacementsNotPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);
            
            if (File.Exists(generatedFileName))
            {
                File.Delete(generatedFileName);
            }

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false, throwExceptions: false);

            // Assert
            Assert.False(actual);
            Assert.False(File.Exists(generatedFileName));
        }

        [Fact]
        public void ReplaceStatic_OneFileCantBeReplaced_ReturnsFalseButReplacementsPerformedOnAllOtherFiles()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            var sourceFilePath = RelativeSourcesPath + "Normal/";

            var sourceFileNames = new List<string>()
            {
                sourceFilePath + "source-resume.txt",
                sourceFilePath + "source-resume.csv",
                sourceFilePath + "source-resume.tsv",
                sourceFilePath + "this-file-doesnt-exist.txt",
                sourceFilePath + "source-resume.docx",
                sourceFilePath + "source-financial-sample.xlsx",
            };

            var outputFileNames = new List<string>()
            {
                RelativeGeneratedFilePath + "output-resume.txt",
                RelativeGeneratedFilePath + "output-resume.csv",
                RelativeGeneratedFilePath + "output-resume.tsv",
                RelativeGeneratedFilePath + "NewDirectory1/NewDirectory2/NewDirectory3/" + "this-file-doesnt-exist.txt",
                RelativeGeneratedFilePath + "output-resume.docx",
                RelativeGeneratedFilePath + "output-financial-sample.xlsx",
            };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            if (File.Exists(RelativeGeneratedFilePath + "this-file-doesnt-exist.txt"))
            {
                File.Delete(RelativeGeneratedFilePath + "this-file-doesnt-exist.txt");
            }

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: false, throwExceptions: false);

            // Assert
            Assert.False(actual);

            foreach (var sourceFile in sourceFiles)
            {
                // verify that no file was created when the source file is invalid, and that all directories that were created are deleted also
                if (Path.GetFileName(sourceFile.SourceFileName) == "this-file-doesnt-exist.txt")
                {
                    Assert.False(File.Exists(RelativeGeneratedFilePath + "NewDirectory1/NewDirectory2/NewDirectory3/" + "this-file-doesnt-exist.txt"));
                    Assert.False(Directory.Exists(RelativeGeneratedFilePath + "NewDirectory1/NewDirectory2/NewDirectory3/"));
                    Assert.False(Directory.Exists(RelativeGeneratedFilePath + "NewDirectory1/NewDirectory2/"));
                    Assert.False(Directory.Exists(RelativeGeneratedFilePath + "NewDirectory1/"));
                    Assert.True(Directory.Exists(RelativeGeneratedFilePath));
                    continue;
                }

                // Compare the generated files to the mock files
                if (Path.GetExtension(sourceFile.OutputFileName) == ".xlsx" ||
                    Path.GetExtension(sourceFile.OutputFileName) == ".docx")
                {
                    Assert.True(FileComparer.FilesAreEqual_OpenXml($"{RelativeMockOutputsPath}Normal/{Path.GetFileName(sourceFile.OutputFileName)}", sourceFile.OutputFileName));
                }
                else
                {
                    Assert.True(FileComparer.FilesAreEqual($"{RelativeMockOutputsPath}Normal/{Path.GetFileName(sourceFile.OutputFileName)}", sourceFile.OutputFileName));
                }
            }

            // Cleanup
            foreach (var sourceFile in sourceFiles)
            {
                if (File.Exists(sourceFile.OutputFileName))
                {
                    File.Delete(sourceFile.OutputFileName);
                }
            }
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume-whole-word.txt")]
        [InlineData("source-resume.csv", "output-resume-whole-word.csv")]
        [InlineData("source-resume.tsv", "output-resume-whole-word.tsv")]
        [InlineData("source-resume.docx", "output-resume-whole-word.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-whole-word.xlsx")]
        public void ReplaceStatic_WholeWord_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "WholeWord/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: true, caseSensitive: false, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume-case-sensitive.txt")]
        [InlineData("source-resume.csv", "output-resume-case-sensitive.csv")]
        [InlineData("source-resume.tsv", "output-resume-case-sensitive.tsv")]
        [InlineData("source-resume.docx", "output-resume-case-sensitive.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-case-sensitive.xlsx")]
        public void ReplaceStatic_CaseSensitive_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "CaseSensitive/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: true, preserveCase: false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume-preserve-case.txt")]
        [InlineData("source-resume.csv", "output-resume-preserve-case.csv")]
        [InlineData("source-resume.tsv", "output-resume-preserve-case.tsv")]
        [InlineData("source-resume.docx", "output-resume-preserve-case.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-preserve-case.xlsx")]
        public void ReplaceStatic_PreserveCase_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "PreserveCase/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: false, caseSensitive: false, preserveCase: true);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt", "output-resume-wwccpc.txt")]
        [InlineData("source-resume.csv", "output-resume-wwccpc.csv")]
        [InlineData("source-resume.tsv", "output-resume-wwccpc.tsv")]
        [InlineData("source-resume.docx", "output-resume-wwccpc.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-wwccpc.xlsx")]
        public void ReplaceStatic_WholeWordCaseSensitiveAndPreserveCase_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "WholeWordCaseSensitiveAndPreserveCase/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, wholeWord: true, caseSensitive: true, preserveCase: true);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-bold.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-bold.xlsx")]
        public void ReplaceStatic_Bold_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(bold: true);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-italics.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-italics.xlsx")]
        public void ReplaceStatic_Italics_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(italics: true);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-underline.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-underline.xlsx")]
        public void ReplaceStatic_Underline_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(underline: true);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-strikethrough.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-strikethrough.xlsx")]
        public void ReplaceStatic_Strikethrough_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(strikethrough: true);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-highlight-red.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-highlight-red.xlsx")]
        public void ReplaceStatic_HighlightRed_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(highlightColor: "#FF0000");

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-text-color-blue.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-text-color-blue.xlsx")]
        public void ReplaceStatic_TextColorBlue_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(textColor: "#0000FF");

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.docx", "output-resume-all-styles.docx")]
        [InlineData("source-financial-sample.xlsx", "output-financial-sample-all-styles.xlsx")]
        public void ReplaceStatic_AllStyling_ReturnsTrueReplacementsPerformed(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "Styling/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(true, true, true, true, "#FF0000", "#0000FF");

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling: styling);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume-split-runs.docx", "output-resume-split-runs.docx")]
        [InlineData("source-financial-sample-split-runs.xlsx", "output-financial-sample-split-runs.xlsx")]
        public void ReplaceStatic_SplitRunsInSources_ReturnsTrueReplacementsPerformedCorrectly(string sourceFileName, string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            sourceFileName = RelativeSourcesPath + "SplitRuns/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "SplitRuns/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false);

            // Assert
            Assert.True(actual);
            Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume-nested-replacements.txt", "output-resume-nested-replacements.txt")]
        [InlineData("source-resume-nested-replacements.docx", "output-resume-nested-replacements.docx")]
        [InlineData("source-financial-sample-nested-replacements.xlsx", "output-financial-sample-nested-replacements.xlsx")]
        public void ReplaceStatic_NestedReplacements_ReturnsTrueReplacementsPerformedCorrectly(string sourceFileName, string outputFileName)
        {
            // If replacements are nested (such as "there" and "her", they are supposed to be handled as follows:
            //      if two matches are nexted, favor the one that starts first
            //          - ex: for "there" and "her", "there" is favored
            //      if two matches start at the same position, favor the longer one
            //          - ex: for "add" and "additional", favor "additional"

            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-nested.csv";
            sourceFileName = RelativeSourcesPath + "NestedReplacements/" + sourceFileName;
            var mockOutputName = RelativeMockOutputsPath + "NestedReplacements/" + outputFileName;
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false);

            // Assert
            Assert.True(actual);

            // Compare the generated file to the mock file
            if (Path.GetExtension(sourceFiles.First().OutputFileName) == ".xlsx" ||
                Path.GetExtension(sourceFiles.First().OutputFileName) == ".docx")
            {
                Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));
            }
            else
            {
                Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));
            }

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Fact]
        public void ReplaceStatic_TextToDocx_ReturnsTrueReplacementsPerformed()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            var sourceFileName = RelativeSourcesPath + "Normal/source-resume.txt";
            var mockOutputName = RelativeMockOutputsPath + "/FileTypeConversion/output-resume.docx";
            var generatedFileName = RelativeGeneratedFilePath + "output-resume.docx";

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false);

            // Assert
            Assert.True(actual);
            Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Fact]
        public void ReplaceStatic_TextToDocxAllStyling_ReturnsTrueReplacementsPerformed()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            var sourceFileName = RelativeSourcesPath + "Normal/source-resume.txt";
            var mockOutputName = RelativeMockOutputsPath + "/FileTypeConversion/output-resume-all-styles.docx";
            var generatedFileName = RelativeGeneratedFilePath + "output-resume-all-styles.docx";

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(true, true, true, true, "#FF0000", "#0000FF");

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling);

            // Assert
            Assert.True(actual);
            Assert.True(FileComparer.FilesAreEqual_OpenXml(mockOutputName, generatedFileName));

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Fact]
        public void ReplaceStatic_DocxToText_ReturnsTrueReplacementsPerformed()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-abbreviations.csv";
            var sourceFileName = RelativeSourcesPath + "Normal/source-resume.docx";
            var mockOutputName = RelativeMockOutputsPath + "/FileTypeConversion/output-resume.txt";
            var generatedFileName = RelativeGeneratedFilePath + "output-resume.txt";

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            var styling = new Styling(true, true, true, true, "#FF0000", "#0000FF");

            // Act
            var actual = QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false, styling);

            // Assert
            Assert.True(actual);
            Assert.True(FileComparer.FilesAreEqual(mockOutputName, generatedFileName));

            // Cleanup
            File.Delete(generatedFileName);
        }

        [Theory]
        [InlineData("source-resume.txt")]
        [InlineData("source-resume.docx")]
        public void ReplaceStatic_AnythingToExcel_ThrowsNotSupportedException(string sourceFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-nested.csv";
            sourceFileName = RelativeSourcesPath + "Normal/" + sourceFileName;
            var generatedFileName = RelativeGeneratedFilePath + "output-resume.xlsx";

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act and Assert
            Assert.Throws<NotSupportedException>(() => QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false));
        }

        [Theory]
        [InlineData("output-financial-sample.txt")]
        [InlineData("output-financial-sample.docx")]
        public void ReplaceStatic_ExcelToAnything_ThrowsNotSupportedException(string outputFileName)
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-nested.csv";
            var sourceFileName = RelativeSourcesPath + "Normal/" + "source-financial-sample.xlsx";
            var generatedFileName = RelativeGeneratedFilePath + outputFileName;

            var sourceFileNames = new List<string>() { sourceFileName };
            var outputFileNames = new List<string>() { generatedFileName };

            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Act and Assert
            Assert.Throws<NotSupportedException>(() => QuickReplace.Replace(replacePhrases, sourceFiles, false, false, false));
        }

        [Theory]
        [InlineData("replacements-abbreviations.csv")]
        [InlineData("replacements-abbreviations.tsv")]
        [InlineData("replacements-abbreviations.xlsx")]
        [InlineData("replacements-abbreviations-pound-delimiter.txt")]
        [InlineData("replacements-abbreviations-semicolon-delimiter.txt")]
        public void ParseReplacements_ValidFile_ReplacementsParsed(string replacementsFileName)
        {
            // Arrange
            replacementsFileName = RelativeReplacementsPath + replacementsFileName;

            // Act
            var replacePhrases = QuickReplace.ParseReplacements(replacementsFileName);

            // Assert
            Assert.Equal("abv-mntnd", replacePhrases["Above-mentioned"]);
            Assert.Equal("Acpt", replacePhrases["Acc,epted"]);
            Assert.Equal("Yr", replacePhrases["Year"]);
        }

        [Fact]
        public void ParseReplacements_InvalidFileType_ThrowsInvalidFileTypeException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "invalid-file-type.tmp";

            // Act and Assert
            Assert.Throws<InvalidFileTypeException>(() => QuickReplace.ParseReplacements(replacementsFileName));
        }

        [Fact]
        public void ParseReplacements_FileDoesntExist_ThrowsFileNotFoundException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "file-doesnt-exist.txt";

            // Act and Assert
            Assert.Throws<FileNotFoundException>(() => QuickReplace.ParseReplacements(replacementsFileName));
        }

        [Fact]
        public void ParseReplacements_EmptyReplacements_ThrowsInvalidOperationException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-empty.csv";

            // Act and Assert
            Assert.Throws<InvalidOperationException>(() => QuickReplace.ParseReplacements(replacementsFileName));
        }

        [Fact]
        public void ParseReplacements_EmptyReplacePhraseItem1_ThrowsInvalidOperationException()
        {
            // Arrange
            var replacementsFileName = RelativeReplacementsPath + "replacements-empty-item1.csv";

            // Act and Assert
            Assert.Throws<InvalidOperationException>(() => QuickReplace.ParseReplacements(replacementsFileName));
        }

        [Fact]
        public void ZipSourceFiles_ValidSourceAndOutputFiles_SourceFilesZipped()
        {
            // Arrange
            var sourceFileNames = new List<string>()
            {
                "source-file-name.txt",
                "source-file-name.csv",
                "source-file-name.tsv",
                "source-file-name.docx",
                "source-file-name.xlsx"
            };

            var outputFileNames = new List<string>()
            {
                "output-file-name.txt",
                "output-file-name.csv",
                "output-file-name.tsv",
                "output-file-name.docx",
                "output-file-name.xlsx"
            };

            // Act
            var sourceFiles = QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames);

            // Assert
            Assert.Equal(5, sourceFiles.Count);
            for (int i = 0; i < sourceFileNames.Count; i++)
            {
                Assert.Equal(sourceFileNames[i], sourceFiles[i].SourceFileName);
                Assert.Equal(outputFileNames[i], sourceFiles[i].OutputFileName);
                Assert.Equal(-1, sourceFiles[i].NumOfReplacements);
            }
        }

        [Fact]
        public void ZipSourceFiles_DifferentSourceAndOutputFileLengths_ThrowsArgumentException()
        {
            // Arrange
            var sourceFileNames = new List<string>() { "source-file-name.txt", "source-file-name.docx" };
            var outputFileNames = new List<string>() { "output-file-name.txt" };

            // Act and Assert
            Assert.Throws<ArgumentException>(() => QuickReplace.ZipSourceFiles(sourceFileNames, outputFileNames));
        }

        [Theory]
        [InlineData("filename.txt")]
        [InlineData("filename.csv")]
        [InlineData("filename.tsv")]
        [InlineData("filename.xlsx")]
        public void IsReplacementFileTypeValid_ValidFileType_ReturnsTrue(string filename)
        {
            // Act
            var actual = QuickReplace.IsReplacementFileTypeValid(filename);

            // Assert
            Assert.True(actual);
        }

        [Theory]
        [InlineData("filename.tmp")]
        [InlineData("filename.docx")]
        public void IsReplacementFileTypeValid_InvalidFileType_ReturnsFalse(string filename)
        {
            // Act
            var actual = QuickReplace.IsReplacementFileTypeValid(filename);

            // Assert
            Assert.False(actual);
        }

        [Fact]
        public void AreSourceFileTypesValid_ValidFileTypes_ReturnsTrue()
        {
            // Arrange
            var sourceFileNames = new List<string>()
            {
                "filename.txt",
                "filename.csv",
                "filename.tsv",
                "filename.xlsx"
            };

            // Act
            var actual = QuickReplace.AreSourceFileTypesValid(sourceFileNames);

            // Assert
            Assert.True(actual);
        }

        [Fact]
        public void AreSourceFileTypesValid_InvalidFileTypes_ReturnsFalse()
        {
            // Arrange
            var sourceFileNames = new List<string>()
            {
                "filename.txt",
                "filename.csv",
                "invalid-filename.tmp",
                "filename.tsv",
                "filename.xlsx"
            };

            // Act
            var actual = QuickReplace.AreSourceFileTypesValid(sourceFileNames);

            // Assert
            Assert.False(actual);
        }

        [Theory]
        [InlineData("filename.txt")]
        [InlineData("filename.csv")]
        [InlineData("filename.tsv")]
        [InlineData("filename.docx")]
        [InlineData("filename.xlsx")]
        public void IsSourceFileTypeValid_ValidFileType_ReturnsTrue(string filename)
        {
            // Act
            var actual = QuickReplace.IsSourceFileTypeValid(filename);

            // Assert
            Assert.True(actual);
        }

        [Fact]
        public void IsSourceFileTypeValid_InvalidFileType_ReturnsFalse()
        {
            // Arrange
            var filename = "invalid-file-type.tmp";

            // Act
            var actual = QuickReplace.IsSourceFileTypeValid(filename);

            // Assert
            Assert.False(actual);
        }

        [Fact]
        public void AreOutputFileTypesValid_ValidFileTypes_ReturnsTrue()
        {
            // Arrange
            var outputFileNames = new List<string>()
            {
                "filename.txt",
                "filename.csv",
                "filename.tsv",
                "filename.xlsx"
            };

            // Act
            var actual = QuickReplace.AreOutputFileTypesValid(outputFileNames);

            // Assert
            Assert.True(actual);
        }

        [Fact]
        public void AreOutputFileTypesValid_InvalidFileTypes_ReturnsFalse()
        {
            // Arrange
            var outputFileNames = new List<string>()
            {
                "filename.txt",
                "filename.csv",
                "invalid-filename.tmp",
                "filename.tsv",
                "filename.xlsx"
            };

            // Act
            var actual = QuickReplace.AreOutputFileTypesValid(outputFileNames);

            // Assert
            Assert.False(actual);
        }

        [Theory]
        [InlineData("filename.txt")]
        [InlineData("filename.csv")]
        [InlineData("filename.tsv")]
        [InlineData("filename.docx")]
        [InlineData("filename.xlsx")]
        public void IsOutputFileTypeValid_ValidFileType_ReturnsTrue(string filename)
        {
            // Act
            var actual = QuickReplace.IsOutputFileTypeValid(filename);

            // Assert
            Assert.True(actual);
        }

        [Fact]
        public void IsOutputFileTypeValid_InvalidFileType_ReturnsFalse()
        {
            // Arrange
            var filename = "invalid-file-type.tmp";

            // Act
            var actual = QuickReplace.IsOutputFileTypeValid(filename);

            // Assert
            Assert.False(actual);
        }

        [Fact]
        public void GenerateMatcher_ValidReplacePhrases_SetsMatcher()
        {
            // Arrange
            var replacePhrases = new Dictionary<string, string>
            {
                {"basic-text", "basic-text1"},
                {"text with whitespace", "text with whitespace 1"},
                {"text,with,commas", "text,with,commas,1"},
                {"text,with\",commas\"and\"quotes,\"", "text,with\",commas\"and\"quotes,\"1"},
                {"text;with;semicolons", "text;with;semicolons1"}
            };

            var sourceFileNames = new List<string>();
            var outputFileNames = new List<string>();

            var QuickReplace = new QuickReplace(replacePhrases, sourceFileNames, outputFileNames);
            var expectedBeforeAct = QuickReplace.IsMatcherCreated();

            // Act
            QuickReplace.GenerateMatcher(true);
            var expectedAfterAct = QuickReplace.IsMatcherCreated();

            // Assert
            Assert.False(expectedBeforeAct);
            Assert.True(expectedAfterAct);
        }

        [Fact]
        public void GenerateMatcher_EmptyReplacePhrases_ThrowsInvalidOperationException()
        {
            // Arrange
            var replacePhrases = new Dictionary<string, string>();

            var sourceFileNames = new List<string>();
            var outputFileNames = new List<string>();

            var QuickReplace = new QuickReplace(replacePhrases, sourceFileNames, outputFileNames);

            // Act and Assert
            Assert.Throws<InvalidOperationException>(() => QuickReplace.GenerateMatcher(true));
        }
        
        [Fact]
        public void ClearMatcher_IsCalled_SetsMatcherToNull()
        {
            // Arrange
            var replacePhrases = new Dictionary<string, string>
            {
                {"basic-text", "basic-text1"},
                {"text with whitespace", "text with whitespace 1"},
                {"text,with,commas", "text,with,commas,1"},
                {"text,with\",commas\"and\"quotes,\"", "text,with\",commas\"and\"quotes,\"1"},
                {"text;with;semicolons", "text;with;semicolons1"}
            };

            var sourceFileNames = new List<string>();
            var outputFileNames = new List<string>();

            var QuickReplace = new QuickReplace(replacePhrases, sourceFileNames, outputFileNames);
            var expected1 = QuickReplace.IsMatcherCreated();

            QuickReplace.GenerateMatcher(true);
            var expected2 = QuickReplace.IsMatcherCreated();

            // Act
            QuickReplace.ClearMatcher();
            var expected3 = QuickReplace.IsMatcherCreated();

            // Assert
            Assert.False(expected1);
            Assert.True(expected2);
            Assert.False(expected3);
        }
    }
}