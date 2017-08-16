using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Syncfusion.XlsIO;
using static GenerateTrikotaPriceLists.DataModule;
using System.IO;
using System.Data;

namespace GenerateTrikotaPriceLists
{
    public static class ExportToExcel
    {
        public static void DoExportToExcel(Client client, DataTable table)
        {
            logger.Info($"Выгрузка прайс-листа в excel для {client.clientDescription}...");

            StringBuilder fileName = new StringBuilder();
            fileName.Append(GetConstant("pricelist-filename"));
            if (GetConstant("fixed-export-path") == "1")
                fileName.Append($"{client.clientDescription}_{client.contractCode}".Replace(' ', '_'));
            fileName.Append(".xml");

            string filePath = Path.Combine(client.exportPath, RemovePathInvalidChars(fileName.ToString(), "_"));
        }

        public static DataTable GetPreparedTable(List<ProductGroup> clientProductGroups, List<Product> clientProducts)
        {
            DataTable table = new DataTable();
            table.Columns.Add("groupCode");
            table.Columns.Add("groupDescription");
            table.Columns.Add("elements", table.GetType());
            table.Columns.Add("children", table.GetType());



            return table;
        }

        public static void AddGroupToPreparedTable(DataTable table, string[] groups, List<ProductGroup> clientProductGroups)
        {

        }
    }
}
