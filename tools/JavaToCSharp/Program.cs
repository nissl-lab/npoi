using System;
using System.Collections.Generic;
using System.Text;
using JavaToCSharp.Configuration;
using System.Configuration;
using JavaToCSharp.Rules;

namespace JavaToCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Missing Argument - file name");
                return;
            }

            RuleEngine re = new RuleEngine();
            re.LoadRules();
            re.Run(args[0]);
        }
    }
}
