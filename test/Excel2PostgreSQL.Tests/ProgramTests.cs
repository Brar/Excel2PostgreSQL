using Excel2PostgreSQL;
using Xunit;

namespace Excel2PostgreSQL.Tests
{
    public class ProgramTests
    {
        [Fact, UseCulture("en")]
        public void Main_WithEmptyArgsArgument_Returns_ErrorMissingExcelFileArgAndPrintsErrorToStdErr()
        {
            using (var console = ConsoleOutput.StartCapturing(false))
            {
                string[] args = { };

                var retVal = (ExitCode)Program.Main(args);

                Assert.Equal(ExitCode.ErrorMissingExcelFileArg, retVal);
                Assert.Contains("Argument missing. Please specify the Excel file to import as Argument.", console.GetStdErrText());
            }
        }

        [Fact, UseCulture("en")]
        public void Main_WithNonexistingFileArg_Returns_ErrorFileDoesNotExistAndPrintsErrorToStdErr()
        {
            const string file = "IAmGone.txt";
            using (var console = ConsoleOutput.StartCapturing(false))
            {
                string[] args = { file };

                var retVal = (ExitCode)Program.Main(args);

                Assert.Equal(ExitCode.ErrorFileDoesNotExist, retVal);
                Assert.Contains($"The file '{file}' doesn't exist.", console.GetStdErrText());
            }
        }
    }
}
