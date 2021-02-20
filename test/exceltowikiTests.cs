using System;
using Xunit;

namespace exceltowiki.Test
{
    using exceltowiki;
    using System.IO;
    using System.Threading;

    public class exceltowikiTests : IClassFixture<TestFileFixture>
    {
        private readonly TestFileFixture _fixture;
        public exceltowikiTests(TestFileFixture fixture)
        {
            _fixture = fixture;
        }
        [Fact]
        public void UsageTest()
        {
            string file = "usage.txt";
            string test = _fixture.GetTempFilePath(file);
            Console.SetError(new StreamWriter(test));
            Program.Main(new string[0]);
            Console.Error.Close();
            Assert.True(_fixture.CompareFiles(_fixture.GetOutputFilePath(file), test));
        }
        [Fact]
        public void HelpTest()
        {
            string file = "help.txt";
            string test = _fixture.GetTempFilePath(file);
            Console.SetError(new StreamWriter(test));
            Program.Main(new string[] { "--help" });
            Console.Error.Close();
            Assert.True(_fixture.CompareFiles(_fixture.GetOutputFilePath(file), test));
        }
        [Fact]
        public void VersionTest()
        {
            string file = "version.txt";
            string test = _fixture.GetTempFilePath(file);
            Console.SetError(new StreamWriter(test));
            Program.Main(new string[] { "--version" });
            Console.Error.Close();
            Assert.True(_fixture.CompareFiles(_fixture.GetOutputFilePath(file), test));
        }
        [Fact]
        public void BasicTest()
        {
            string file = "basic.wiki";
            string test = _fixture.GetTempFilePath(file);
            Console.SetOut(new StreamWriter(test));
            Program.Main(new string[] { @"input\API-metadata.xlsx" });
            Console.Out.Close();
            Assert.True(_fixture.CompareFiles(_fixture.GetOutputFilePath(file), test));
        }
    }
}
