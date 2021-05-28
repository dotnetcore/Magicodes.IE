using System;
using System.Reflection;

namespace Magicodes.IE.Tools
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Hello Magicodes.IE!");

            if (args.Length == 0)
            {
                var versionString = Assembly.GetEntryAssembly()
                                        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                                        .InformationalVersion
                                        .ToString();

                Console.WriteLine($"mie v{versionString}");
                Console.WriteLine("-------------");
                Console.WriteLine("\nGithub:");
                Console.WriteLine("  https://github.com/dotnetcore/Magicodes.IE");
                return;
            }
        }
    }
}