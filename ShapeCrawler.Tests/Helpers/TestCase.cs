using System.Globalization;

namespace ShapeCrawler.Tests.Helpers
{
    public class TestCase<T1, T2>
    {
        private readonly int testCaseNumber;
        public T1 Param1 { get; }
        public T2 Param2 { get; }

        public TestCase(int testCaseNumber, T1 param1, T2 param2)
        {
            this.testCaseNumber = testCaseNumber;
            this.Param1 = param1;
            this.Param2 = param2;
        }

        public override string ToString()
        {
            return this.testCaseNumber.ToString(NumberFormatInfo.CurrentInfo);
        }
    }
}