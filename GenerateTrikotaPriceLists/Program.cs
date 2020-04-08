using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static GenerateTrikotaPriceLists.DataModule;

namespace GenerateTrikotaPriceLists
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                LoadConstants(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_ОбщаяИнформация.txt"));
                LoadMatrix(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_ТоварнаяМатрица.txt"));
                LoadAllMatrixProducts(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_НоменклатураТоварнойМатрицы.txt"));
                LoadProducts(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_Остатки.txt"));
                LoadProductPrices(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_Цены.txt"));
                LoadClients(Path.Combine(Environment.CurrentDirectory, "ВыгрузкаПрайсЛистов_Контрагенты.txt"));

                FillProductGroups(products, productGroups);

                ExportPriceLists();
            }
            catch (Exception exception)
            {
                logger.Fatal(exception);
            }
        }
    }
}
