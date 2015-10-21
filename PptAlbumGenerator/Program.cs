using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptAlbumGenerator
{
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine(Prompts.HeaderText);
            Console.WriteLine();
            if (args.Length < 1)
            {
                ShowHelp();
                return 1;
            }
            using (var reader = new StreamReader(args[0]))
            {
                var generator = new AlbumGenerator(reader)
                {
                    DefaultWorkPath = Path.GetDirectoryName(args[0])
                };
                generator.Generate();
            }
            return 0;
        }

        static void ShowHelp()
        {
            Console.WriteLine(Prompts.HelpText);
        }
    }
}
