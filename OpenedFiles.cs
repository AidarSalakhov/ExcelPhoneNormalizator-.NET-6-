using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPhoneNormalizator
{
    internal class OpenedFiles
    {
        public int _index { get; set; }
        public string _fileName { get; set; }
        public string _projectName { get; set; }
        public int _leadsCont { get; set; }


        public void Print()
        {
            Console.WriteLine($"{_index}) {_fileName}");
            Console.WriteLine($"Название проекта: {_projectName}");
            Console.WriteLine($"Количество заявок: {_leadsCont}\n");
        }

    }
}
