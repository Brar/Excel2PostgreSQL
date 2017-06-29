using Excel2PostgreSQL;
using Xunit;

namespace Excel2PostgreSQL.Tests
{
    public class ProgramTests
    {
        [Fact, UseCulture("en")]
        public void Main_WithEmptyArgsArgument_Returns_1AndPrintsErrorToStdErr()
        {
            using (var console = ConsoleOutput.StartCapturing(false))
            {
                string[] args = { };

                var retVal = Program.Main(args);

                Assert.Equal(1, retVal);
                Assert.Contains("Argument missing. Please specify the Excel file to import as Argument.", console.GetStdErrText());
            }
        }
    }
}
