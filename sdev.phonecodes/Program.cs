using CsvHelper;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sdev.phonecodes
{
    public class Operator
    {
        public Int64 From { get; set; }
        public Int64 To { get; set; }
        public string Name { get; set; }
        public string Region { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Поиск операторов и телефонов по базе DEF. (c) 2015 Игорь Подсекин. wonderu@wonderu.com");
            if (args.Length != 3)
            {
                Console.WriteLine("Необходимо указать параметры запуска: <полный путь до Excel-файла с номерами> <номер колонки с телефонными номерами> <номер колонки, с которой будут записываться названия операторов и регионы>");
                return;
            }
            var fileInfoIn = new FileInfo(args[0]);
            if (!fileInfoIn.Exists)
            {
                Console.WriteLine("Указанный файл не существует");
                return;
            }

            int numIndex;


            if (!(Int32.TryParse(args[1], out numIndex) && numIndex > 0))
            {
                Console.WriteLine("Номер колонки с телефонными номерами должен быть целым числом больше нуля");
                return;
            }

            int resultIndex;
            if (!(Int32.TryParse(args[2], out resultIndex) && resultIndex > 0))
            {
                Console.WriteLine("номер колонки, с которой будут записываться названия операторов и регионы, должен быть целым числом больше нуля");
                return;
            }

            SortedList<Int64, Operator> numbers = new SortedList<Int64, Operator>();

            var dir = Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]);
            
            var files = Directory.EnumerateFiles(dir + "\\codes", "*.csv");
            foreach (var file in files)
            {
                using (var reader = File.OpenText(file))
                {
                    var csv = new CsvReader(reader, new CsvHelper.Configuration.CsvConfiguration()
                    {
                        Delimiter = ";"
                    });

                    while (csv.Read())
                    {
                        var oper = new Operator()
                        {
                            From = csv.GetField<Int64>(0) * 10000000 + csv.GetField<Int64>(1),
                            To = csv.GetField<Int64>(0) * 10000000 + csv.GetField<Int64>(2),
                            Name = csv.GetField<string>(4).Trim(),
                            Region = csv.GetField<string>(5).Trim(),
                        };

                        numbers.Add(oper.From, oper);
                    }
                }
            }

            using (var package = new ExcelPackage(fileInfoIn))
            {
                ExcelWorkbook workBook = package.Workbook;
                if (workBook != null)
                {
                    if (workBook.Worksheets.Count > 0)
                    {
                        ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

                        for (int i = 2; i <= currentWorksheet.Cells.Rows && currentWorksheet.Cells[i, numIndex].Value != null; i++)
                        {
                            Int64 inNumber;
                            if (Int64.TryParse(currentWorksheet.Cells[i, numIndex].Value.ToString(), out inNumber))
                            {
                                inNumber = inNumber % 10000000000;
                                var oper = numbers.FirstOrDefault(k => k.Value.From <= inNumber && k.Value.To >= inNumber);
                                if (oper.Value != null)
                                {
                                    currentWorksheet.Cells[i, resultIndex].Value = oper.Value.Name;
                                    currentWorksheet.Cells[i, resultIndex + 1].Value = oper.Value.Region;
                                }
                            }
                        }
                    }
                }
                package.Save();
            }
        }
    }
}
