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
        static void Main(string[] args)
        {
            using (var reader = new StreamReader(args[0]))
            {
                var generator = new AlbumGenerator(reader);
                generator.Generate();
            }
        }
    }
}
