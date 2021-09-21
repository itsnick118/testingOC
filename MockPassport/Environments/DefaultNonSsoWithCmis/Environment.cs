using System;
using System.IO;

namespace MockPassport.Environments.DefaultNonSsoWithCmis
{
    public class Environment : IEnvironment
    {
        public Uri BaseUri => new Uri("https://gts-ey-qa/Passport");
        public string Username => "admin";
        public string Password => "datacert";
        public string Name => "DefaultNonSsoWithCmis";
        public DirectoryInfo BaseFilePath => new DirectoryInfo($"..\\..\\Environments\\{Name}");
        public int ModelMatter => 17;
        public int ModelInvoiceHeader => 44;
    }
}