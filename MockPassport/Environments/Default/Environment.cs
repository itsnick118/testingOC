using System;
using System.IO;
using MockPassport.Mappings;

namespace MockPassport.Environments.Default
{
    public class Environment : IEnvironment
    {
        public Uri BaseUri => new Uri("https://gts-ey-qa/Passport");
        public string Username => "admin";
        public string Password => "datacert";
        public string Name => "Default";
        public DirectoryInfo BaseFilePath => new DirectoryInfo("..\\..");
        public int ModelMatter => 8;
        public int ModelInvoiceHeader => 44;
    }
}