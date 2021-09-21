namespace MockPassport
{
    public class CommandLineArguments
    {
        public const string DefaultEnvironment = "DefaultNonSsoWithCmis";

        public string Environment { get; set; }

        public bool Update { get; set; }

        public string Record { get; set; }
    }
}