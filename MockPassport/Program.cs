using System;
using Microsoft.Test.CommandLineParsing;

namespace MockPassport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var commandLineArguments = new CommandLineArguments();

            try
            {
                commandLineArguments.ParseArguments(args);
            }
            catch (Exception exception)
            {
                Console.WriteLine();
                Console.WriteLine(exception.Message);
                Console.WriteLine();
                Console.WriteLine("Usage: MockPassport [/update] ");
                Console.WriteLine("                    [/environment=<environmentname>]");
                return;
            }
            
            var environment = string.IsNullOrEmpty(commandLineArguments.Environment)
                ? CommandLineArguments.DefaultEnvironment
                : commandLineArguments.Environment;

            if (commandLineArguments.Update)
            {
                MockEnvironment.Update(environment);
            }
            else
            {
                Socket.Start();
                MockEnvironment.Start(environment);
            }
        }
    }
}
