using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Syncfusion.XlsIO;
using static GenerateTrikotaPriceLists.DataModule;
using System.IO;
using System.Data;
using System.Drawing;

namespace GenerateTrikotaPriceLists
{
    public static class ExportToExcel
    {
        public static void DoExportToExcel(Client client, List<Product> clientProducts, DataTable table)
        {
            logger.Info($"Выгрузка прайс-листа в excel для {client.clientDescription}...");

            StringBuilder fileName = new StringBuilder();
            fileName.Append(GetConstant("pricelist-filename"));
            if (GetConstant("fixed-export-path") == "1")
                fileName.Append($"{client.clientDescription}_{client.contractCode}".Replace(' ', '_'));
            fileName.Append(".xls");

            string filePath = Path.Combine(client.exportPath, RemovePathInvalidChars(fileName.ToString(), "_"));

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication excelApplication = excelEngine.Excel;
                excelApplication.DefaultVersion = ExcelVersion.Excel97to2003;

                IWorkbook workbook = excelApplication.Workbooks.Create(new string[] { "price-list" });
                IWorksheet worksheet = workbook.Worksheets[0];

                int columnsCount = 6;
                if (client.isAppendClientCodeExcel)
                    columnsCount += 2;
                if (client.isExportProductArticle)
                    columnsCount++;
                if (client.isExportProductComment)
                    columnsCount++;

                int index;

                AddExcelStyle(workbook, "Arial14BoldBorders", "Arial", 14, true, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignCenter, true, true, true, true);
                AddExcelStyle(workbook, "Arial14BoldLeft", "Arial", 14, true, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, false, false, false, false);
                AddExcelStyle(workbook, "Arial10Left", "Arial", 10, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, false, false, false, false);
                AddExcelStyle(workbook, "Arial10BoldBorders", "Arial", 10, true, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignCenter, true, true, true, true);
                AddExcelStyle(workbook, "ArialGroup1", "Arial", 12, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.Gold);
                AddExcelStyle(workbook, "ArialGroup2", "Arial", 11, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.Yellow);
                AddExcelStyle(workbook, "ArialGroup3", "Arial", 10, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.LightYellow);
                AddExcelStyle(workbook, "ArialGroup4", "Arial", 9, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.FloralWhite);

                if (!String.IsNullOrWhiteSpace(GetConstant("global-comment")))
                    AddTextToExcelCell(workbook, worksheet, GetConstant("global-comment"), 1, 1, 1, columnsCount, "Arial14BoldBorders");

                AddTextToExcelCell(workbook, worksheet, $"Прайс-лист {GetConstant("company-name")} от {DateTime.Now.ToString("dd.MM.yyyy")}", 3, 1, 3, columnsCount, "Arial14BoldLeft");

                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"Адрес: {GetConstant("company-address")}");
                sb.AppendLine($"Телефоны: {GetConstant("company-phone")}");
                sb.AppendLine($"Сайт: {GetConstant("company-web")}");
                sb.Append($"E-mail: {GetConstant("e-mail")}");
                AddTextToExcelCell(workbook, worksheet, sb.ToString(), 4, 1, 4, columnsCount, "Arial10Left");

                sb.Clear();
                sb.Append("Контрагент: ");
                if (client.isAppendClientCodeExcel)
                    sb.Append($"[{client.clientCode}] ");
                sb.Append(client.clientDescription);
                AddTextToExcelCell(workbook, worksheet, sb.ToString(), 6, 1, 6, columnsCount, "Arial10Left");

                index = 1;
                AddTextToExcelCell(workbook, worksheet, "№", 8, index, 8, index++, "Arial10BoldBorders");
                if (client.isAppendClientCodeExcel)
                    AddTextToExcelCell(workbook, worksheet, "Код", 8, index, 8, index++, "Arial10BoldBorders");
                if (client.isExportProductArticle)
                    AddTextToExcelCell(workbook, worksheet, "Артикул", 8, index, 8, index++, "Arial10BoldBorders");
                AddTextToExcelCell(workbook, worksheet, "Номенклатура", 8, index, 8, index++, "Arial10BoldBorders");
                AddTextToExcelCell(workbook, worksheet, "Ед.изм.", 8, index, 8, index++, "Arial10BoldBorders");
                AddTextToExcelCell(workbook, worksheet, "Кол-во\nв упак.", 8, index, 8, index++, "Arial10BoldBorders");
                AddTextToExcelCell(workbook, worksheet, "Остаток", 8, index, 8, index++, "Arial10BoldBorders");
                AddTextToExcelCell(workbook, worksheet, "Цена", 8, index, 8, index++, "Arial10BoldBorders");
                if (client.isExportProductComment)
                    AddTextToExcelCell(workbook, worksheet, "Комментарий", 8, index, 8, index++, "Arial10BoldBorders");
                if (client.isAppendClientCodeExcel)
                    AddTextToExcelCell(workbook, worksheet, "ЗАКАЗ", 8, index, 8, index++, "Arial10BoldBorders");

                WriteTable(workbook, worksheet, table, client, columnsCount, 1);

                workbook.SaveAs(filePath);

                workbook.Close();
            }
        }

        public static void WriteTable(IWorkbook workbook, IWorksheet worksheet, DataTable table, Client client, int columnsCount, int iLevel)
        {
            if (table.Rows.Count == 0)
                return;

            foreach (DataRow row in table.Rows)
            {

            }
        }

        public static void AddTextToExcelCell(IWorkbook workbook, IWorksheet worksheet, string value, int row, int col, int lastRow, int lastCol, string styleName)
        {
            var range = worksheet.Range[row, col, lastRow, lastCol];
            range.CellStyle = workbook.Styles[styleName];
            range.Text = value;
            if (row != lastRow || col != lastCol)
                range.Merge(false);
        }

        public static void AddExcelStyle(IWorkbook workbook, string name, string fontName, int fontSize, bool isFontBold,
            ExcelVAlign vAlign, ExcelHAlign hAlign, bool bLeft, bool bRight, bool bTop, bool bBottom, Color? color = null)
        {
            IStyle style = workbook.Styles.Add(name);

            style.Font.FontName = fontName;
            style.Font.Size = fontSize;
            style.Font.Bold = isFontBold;

            style.VerticalAlignment = vAlign;
            style.HorizontalAlignment = hAlign;

            style.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = bLeft ? ExcelLineStyle.Thin : ExcelLineStyle.None;
            style.Borders[ExcelBordersIndex.EdgeRight].LineStyle = bRight ? ExcelLineStyle.Thin : ExcelLineStyle.None;
            style.Borders[ExcelBordersIndex.EdgeTop].LineStyle = bTop ? ExcelLineStyle.Thin : ExcelLineStyle.None;
            style.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = bBottom ? ExcelLineStyle.Thin : ExcelLineStyle.None;

            if (color != null)
                style.Color = (Color)color;
        }

        public static DataTable GetPreparedTable(List<ProductGroup> clientProductGroups)
        {
            DataTable table = new DataTable();
            table.Columns.Add("groupCode");
            table.Columns.Add("groupDescription");
            table.Columns.Add("level");
            table.Columns.Add("children", table.GetType());

            foreach (ProductGroup group in clientProductGroups)
            {
                string[] levels = group.sLevel.Split('.');
                string level = "";
                DataTable children = table;
                for (int i = 0; i < levels.Count(); i ++)
                {
                    level = $"{level}{(i == 0 ? "" : ".")}{levels[i]}";
                    children = AddGroupToPreparedTable(children, level, clientProductGroups);
                }
            }

            return table;
        }

        public static DataTable AddGroupToPreparedTable(DataTable table, string level, List<ProductGroup> clientProductGroups)
        {
            if (table == null)
                return null;

            var group = clientProductGroups.Where(w => w.sLevel == level).FirstOrDefault();
            if (group == null)
                return null;

            DataRow row = table.Select($"level = '{level}'").FirstOrDefault();
            if (row == null)
            {
                DataTable childrenTable = new DataTable();
                childrenTable.Columns.Add("groupCode");
                childrenTable.Columns.Add("groupDescription");
                childrenTable.Columns.Add("level");
                childrenTable.Columns.Add("children", childrenTable.GetType());

                row = table.NewRow();
                row["groupCode"] = group.code;
                row["groupDescription"] = group.description;
                row["level"] = level;
                row["children"] = childrenTable;

                table.Rows.Add(row);
            }

            return (DataTable)row["children"];
        }
    }
}
