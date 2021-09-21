using System;
using System.IO;
using MockPassport.Environments.Default;
using MockPassport.Mappings;

namespace MockPassport
{
    public interface IEnvironment
    {
        Uri BaseUri { get; }
        string Username { get; }
        string Password { get; }
        string Name { get; }
        DirectoryInfo BaseFilePath { get; }
        int ModelMatter { get; }
        int ModelInvoiceHeader { get; }
    }
}