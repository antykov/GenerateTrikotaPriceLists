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
        public static IWorkbook workbook;
        public static IWorksheet worksheet;

        public static int columnsCount;
        public static int columnNumber, rowNumber;
        public static int priceListNumber;

        public static void DoExportToExcel(Client client, List<Product> clientProducts, DataTable table)
        {
            logger.Info($"Выгрузка прайс-листа в Excel для {client.clientDescription}...");

            StringBuilder fileName = new StringBuilder();
            fileName.Append(GetConstant("pricelist-filename"));
            if (GetConstant("fixed-export-path") == "1")
                fileName.Append($"{client.clientDescription}_{client.contractCode}".Replace(' ', '_'));
            fileName.Append(".xls");

            string filePath = Path.Combine(StrToBoolDef(GetConstant("use-current-directory"), false) ? Environment.CurrentDirectory : client.exportPath, RemovePathInvalidChars(fileName.ToString(), "_"));

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication excelApplication = excelEngine.Excel;
                excelApplication.DefaultVersion = ExcelVersion.Excel97to2003;

                workbook = excelApplication.Workbooks.Create(new string[] { "Прайс-лист" });
                worksheet = workbook.Worksheets[0];

                columnsCount = 6;
                if (client.exportToEXCELVariant == 1)
                {
                    if (client.isAppendClientCodeExcel)
                        columnsCount += 2;
                    if (client.isExportProductArticle)
                        columnsCount++;
                    if (client.isExportProductComment)
                        columnsCount++;
                } else
                {
                    if (client.isAppendClientCodeExcel)
                        columnsCount++;
                }

                rowNumber = 1;

                AddExcelStyle("Arial14BoldCenterBorders", "Arial", 14, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignCenter, true, true, true, true);
                AddExcelStyle("Arial14BoldLeft", "Arial", 14, true, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, false, false, false, false);
                AddExcelStyle("Arial10Left", "Arial", 10, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, false, false, false, false);
                AddExcelStyle("Arial10BoldCenterBorders", "Arial", 10, true, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignCenter, true, true, true, true);
                AddExcelStyle("Arial10LeftBorders", "Arial", 10, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, true, true, true, true);
                AddExcelStyle("Arial10RightBorders", "Arial", 10, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignRight, true, true, true, true);
                AddExcelStyle("Arial10CenterBorders", "Arial", 10, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignCenter, true, true, true, true);
                AddExcelStyle("Arial7LeftBorders", "Arial", 7, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignLeft, true, true, true, true);
                AddExcelStyle("Arial7RightBorders", "Arial", 7, false, ExcelVAlign.VAlignCenter, ExcelHAlign.HAlignRight, true, true, true, true);
                AddExcelStyle("ArialGroup1", "Arial", 12, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.Gold);
                AddExcelStyle("ArialGroup2", "Arial", 11, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.Yellow);
                AddExcelStyle("ArialGroup3", "Arial", 10, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.LightYellow);
                AddExcelStyle("ArialGroup4", "Arial", 9, true, ExcelVAlign.VAlignTop, ExcelHAlign.HAlignLeft, true, true, true, true, Color.FloralWhite);

                if (!String.IsNullOrWhiteSpace(GetConstant("global-comment")))
                {
                    AddTextToExcelCell(GetConstant("global-comment"), rowNumber, 1, rowNumber, columnsCount, "Arial14BoldCenterBorders", false);
                    worksheet.SetRowHeight(rowNumber, worksheet.GetRowHeight(rowNumber) / 2.5);
                    rowNumber++;
                }

                int firstAutoFitRow = rowNumber;

                rowNumber++;
                AddTextToExcelCell($"Прайс-лист {GetConstant("company-name")} от {DateTime.Now.ToString("dd.MM.yyyy")}", rowNumber, 1, rowNumber++, columnsCount, "Arial14BoldLeft");
                AddTextToExcelCell($"Адрес: {GetConstant("company-address")}", rowNumber, 1, rowNumber++, columnsCount, "Arial10Left");
                AddTextToExcelCell($"Телефоны: {GetConstant("company-phone")}", rowNumber, 1, rowNumber++, columnsCount, "Arial10Left");
                AddTextToExcelCell($"Сайт: {GetConstant("company-web")}", rowNumber, 1, rowNumber++, columnsCount, "Arial10Left");
                AddTextToExcelCell($"E-mail: {GetConstant("company-e-mail")}", rowNumber, 1, rowNumber++, columnsCount, "Arial10Left");

                StringBuilder sb = new StringBuilder();
                sb.Append("Контрагент: ");
                if (client.isAppendClientCodeExcel)
                    sb.Append($"[{client.clientCode}] ");
                sb.Append(client.clientDescription);
                rowNumber++;
                AddTextToExcelCell(sb.ToString(), rowNumber, 1, rowNumber++, columnsCount, "Arial10Left");

                rowNumber++;
                columnNumber = 1;
                if (client.exportToEXCELVariant == 1)
                {
                    AddTextToExcelCell("№", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 7);
                    if (client.isAppendClientCodeExcel)
                    {
                        AddTextToExcelCell("Код", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                        worksheet.ShowColumn(columnNumber++, false);
                    }
                    if (client.isExportProductArticle)
                    {
                        AddTextToExcelCell("Артикул", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                        worksheet.SetColumnWidth(columnNumber++, 10);
                    }
                    AddTextToExcelCell("Номенклатура", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 60);
                    AddTextToExcelCell("Ед.изм.", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 7);
                    AddTextToExcelCell("Кол-во\nв упак.", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 7);
                    AddTextToExcelCell("Остаток", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 10);
                    AddTextToExcelCell("Цена", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 10);
                    if (client.isExportProductComment)
                    {
                        AddTextToExcelCell("Комментарий", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                        worksheet.SetColumnWidth(columnNumber++, 30);
                    }
                } else
                {
                    AddTextToExcelCell("Код", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 12);
                    AddTextToExcelCell("Бренд", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 25);
                    AddTextToExcelCell("Наименование", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 80);
                    AddTextToExcelCell("Остаток", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 10);
                    AddTextToExcelCell("Цена", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 10);
                    AddTextToExcelCell("Штрихкод", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 17);
                }
                if (client.isAppendClientCodeExcel)
                {
                    AddTextToExcelCell("ЗАКАЗ", rowNumber, columnNumber, rowNumber, columnNumber, "Arial10BoldCenterBorders");
                    worksheet.SetColumnWidth(columnNumber++, 10);
                }

                rowNumber++;
                priceListNumber = 1;

                if (client.exportToEXCELVariant == 1)
                    WriteTable_Excel_1(table, client, clientProducts, 1);
                else
                    WriteTable_Excel_2(table, client, clientProducts, 1);

                workbook.SaveAs(filePath);

                workbook.Close();
            }  
        }

        public static void WriteTable_Excel_1(DataTable table, Client client, List<Product> clientProducts, int iLevel)
        {
            if (table.Rows.Count == 0)
                return;

            foreach (DataRow row in table.Rows)
            {
                AddTextToExcelCell((string)row["groupDescription"], rowNumber, 1, rowNumber++, columnsCount, $"ArialGroup{Math.Min(iLevel, 4)}");

                int firstGroupingRow = rowNumber;

                WriteTable_Excel_1((DataTable)row["children"], client, clientProducts, iLevel + 1);

                foreach (Product product in clientProducts.Where(w => w.level == (string)row["level"]))
                {
                    columnNumber = 1;
                    AddTextToExcelCell($"{priceListNumber++}", rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10RightBorders");
                    if (client.isAppendClientCodeExcel)
                        AddTextToExcelCell(product.code, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial7LeftBorders");
                    if (client.isExportProductArticle)
                        AddTextToExcelCell(product.article, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial7LeftBorders");
                    AddTextToExcelCell(product.description, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");
                    AddTextToExcelCell(product.unit, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10CenterBorders");
                    AddTextToExcelCell(product.pack, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10CenterBorders");
                    AddTextToExcelCell(product.quantity, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10RightBorders");
                    AddTextToExcelCell(product.price.ToString("0.00"), rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10RightBorders");
                    if (client.isExportProductComment)
                        AddTextToExcelCell(product.comment, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial7LeftBorders");
                    if (client.isAppendClientCodeExcel)
                        AddTextToExcelCell("", rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");

                    rowNumber++;
                }

                worksheet.Range[firstGroupingRow, 1, rowNumber - 1, columnsCount].Group(ExcelGroupBy.ByRows, false);
            }
        }

        public static void WriteTable_Excel_2(DataTable table, Client client, List<Product> clientProducts, int iLevel)
        {
            if (table.Rows.Count == 0)
                return;

            foreach (DataRow row in table.Rows)
            {
                AddTextToExcelCell((string)row["groupDescription"], rowNumber, 1, rowNumber++, columnsCount, $"ArialGroup{Math.Min(iLevel, 4)}");

                int firstGroupingRow = rowNumber;

                WriteTable_Excel_2((DataTable)row["children"], client, clientProducts, iLevel + 1);
                
                foreach (Product product in clientProducts.Where(w => w.level == (string)row["level"]))
                {
                    columnNumber = 1;
                    AddTextToExcelCell(product.code, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");
                    AddTextToExcelCell(product.brand, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");
                    AddTextToExcelCell(product.description, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");
                    AddTextToExcelCell(product.quantity, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10RightBorders");
                    AddTextToExcelCell(product.price.ToString("0.00"), rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10RightBorders");
                    AddTextToExcelCell(product.barcode, rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");
                    if (client.isAppendClientCodeExcel)
                        AddTextToExcelCell("", rowNumber, columnNumber, rowNumber, columnNumber++, "Arial10LeftBorders");

                    rowNumber++;
                }

                worksheet.Range[firstGroupingRow, 1, rowNumber - 1, columnsCount].Group(ExcelGroupBy.ByRows, false);
            }
        }

        public static void AddTextToExcelCell(string value, int row, int col, int lastRow, int lastCol, string styleName, bool isAutofitMergedRow = true)
        {
            var range = worksheet.Range[row, col, lastRow, lastCol];
            range.CellStyle = workbook.Styles[styleName];
            range.Text = value;
            if (row != lastRow || col != lastCol)
            {
                range.Merge(false);
                if (isAutofitMergedRow)
                    worksheet.AutofitRow(row);
            } 
        }

        public static void AddExcelStyle(string name, string fontName, int fontSize, bool isFontBold,
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

            style.WrapText = true;
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
