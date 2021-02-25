using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace exceltowiki.Test
{
    public class FileFixture : IDisposable
    {
        public string TestDir { get; }
        private bool _deleteOnDispose = true;
        public FileFixture()
        {
            System.Diagnostics.Trace.WriteLine("In file fixture constructor");
            TestDir = CreateTestDirectory();

        }
        public virtual void Dispose()
        {
            if(_deleteOnDispose && Directory.Exists(TestDir))
            {
                Directory.Delete(TestDir, true);
            }
        }
        protected string CreateTestDirectory()
        {
            string testbasedir = Path.GetTempPath();
            string scratchpath = Environment.GetEnvironmentVariable("SCRATCH");
            if(!String.IsNullOrEmpty(scratchpath))
            {
                testbasedir = scratchpath;
                _deleteOnDispose = false;
            }

            string foldername = Assembly.GetExecutingAssembly().GetName().Name;
            string testdir = Path.Combine(testbasedir, foldername);
            if (Directory.Exists(testdir))
            {
                foreach(string path in Directory.GetFiles(testdir))
                {
                    File.Delete(path);
                }
            }
            else
            {
                Directory.CreateDirectory(testdir);
            }
            return testdir;
        }
        protected void CreateFile(string path)
        {
            using (FileStream fs = File.Create(path))
            {

            }
        }
    }
}
