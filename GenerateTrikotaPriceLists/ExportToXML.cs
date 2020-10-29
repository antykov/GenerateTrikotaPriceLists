using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using static GenerateTrikotaPriceLists.DataModule;

namespace GenerateTrikotaPriceLists
{
    public static class ExportToXML
    {
        public static void DoExportToXML(Client client, List<ProductGroup> clientProductGroups, List<Product> clientProducts)
        {
            logger.Info($"Выгрузка прайс-листа в Xml для {client.clientDescription}...");

            StringBuilder fileName = new StringBuilder();
            fileName.Append(GetConstant("pricelist-filename"));
            if (GetConstant("fixed-export-path") == "1")
                fileName.Append($"{client.clientDescription}_{client.contractCode}".Replace(' ', '_'));
            fileName.Append(".xml");

            string filePath = Path.Combine(StrToBoolDef(GetConstant("use-current-directory"), false) ? Environment.CurrentDirectory : client.exportPath, RemovePathInvalidChars(fileName.ToString(), "_"));

            try
            {
                List<RetailRest> clientRetailRests = null;
                if (!String.IsNullOrWhiteSpace(client.storehouseCodeForMatrixFilter))
                    clientRetailRests = retailRests.Where(w => w.storeCode == client.storehouseCodeForMatrixFilter).ToList<RetailRest>();
                bool useRetailRests = (clientRetailRests?.Count ?? 0) > 0 ? true : false;

                XmlWriterSettings xmlSettings = new XmlWriterSettings();
                xmlSettings.Encoding = Encoding.GetEncoding(1251);
                xmlSettings.Indent = true;

                using (XmlWriter writer = XmlWriter.Create(filePath, xmlSettings))
                {
                    writer.WriteStartDocument();

                    writer.WriteStartElement("PRICE-LIST");
                    writer.WriteAttributeString("company", GetConstant("company-name"));
                    writer.WriteAttributeString("e-mail", GetConstant("company-e-mail"));
                    writer.WriteAttributeString("date", DateTime.Now.ToString("dd/MM/yyyy HH:mm", System.Globalization.CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("client_code", client.clientCode + (client.isExportByContract ? $"@{client.contractCode}" : ""));
                    writer.WriteAttributeString("client_description", client.clientDescription);

                    string globalComment = GetConstant("global-comment");
                    if (!String.IsNullOrWhiteSpace(globalComment))
                    {
                        writer.WriteStartElement("global_comment");
                        writer.WriteCData(globalComment);
                        writer.WriteEndElement();
                    }

                    WriteCurrencies(writer);
                    WriteFields(writer, client, useRetailRests);
                    WriteGroups(writer, clientProductGroups);
                    WriteData(writer, client, clientProducts, useRetailRests, clientRetailRests);

                    writer.WriteEndElement();

                    writer.WriteEndDocument();
                }
            }
            catch (Exception exception)
            {
                logger.Fatal(exception);
            }
        }

        public static void WriteCurrencies(XmlWriter writer)
        {
            writer.WriteStartElement("currencies");

            writer.WriteStartElement("currency");
            WriteXmlValue(writer, "code", "643");
            WriteXmlValue(writer, "fullname", "Российский рубль");
            WriteXmlValue(writer, "name", "руб.");
            WriteXmlValue(writer, "rate", "1");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        public static void WriteFields(XmlWriter writer, Client client, bool useRetailRests)
        {
            writer.WriteStartElement("fields");

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "code");
            WriteXmlValue(writer, "description", "Код");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "10");
            WriteXmlValue(writer, "align", "-1");
            writer.WriteEndElement();

            if (client.isExportProductArticle)
            {
                writer.WriteStartElement("field");
                WriteXmlValue(writer, "name", "n_article");
                WriteXmlValue(writer, "description", "Артикул");
                WriteXmlValue(writer, "type", "string");
                WriteXmlValue(writer, "length", "25");
                WriteXmlValue(writer, "align", "-1");
                writer.WriteEndElement();
            }

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "description");
            WriteXmlValue(writer, "description", "Наименование");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "100");
            WriteXmlValue(writer, "align", "-1");
            writer.WriteEndElement();

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "unit");
            WriteXmlValue(writer, "description", "Ед. изм.");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "10");
            WriteXmlValue(writer, "align", "0");
            writer.WriteEndElement();

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "quantity");
            WriteXmlValue(writer, "description", "Количество");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "15");
            WriteXmlValue(writer, "align", "1");
            writer.WriteEndElement();

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "pack_coefficient");
            WriteXmlValue(writer, "description", "Упаковка");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "15");
            WriteXmlValue(writer, "align", "1");
            writer.WriteEndElement();

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "price");
            WriteXmlValue(writer, "description", "Цена");
            WriteXmlValue(writer, "type", "double");
            WriteXmlValue(writer, "length", "15");
            WriteXmlValue(writer, "precision", "3");
            WriteXmlValue(writer, "align", "1");
            writer.WriteEndElement();

            writer.WriteStartElement("field");
            WriteXmlValue(writer, "name", "currency");
            WriteXmlValue(writer, "description", "Валюта");
            WriteXmlValue(writer, "type", "string");
            WriteXmlValue(writer, "length", "5");
            WriteXmlValue(writer, "align", "0");
            writer.WriteEndElement();

            if (client.isExportByProductMatrix)
            {
                writer.WriteStartElement("field");
                WriteXmlValue(writer, "name", "recommendQuantity");
                WriteXmlValue(writer, "description", "Рек. количество");
                WriteXmlValue(writer, "type", "string");
                WriteXmlValue(writer, "length", "15");
                WriteXmlValue(writer, "align", "1");
                writer.WriteEndElement();
            }

            if (useRetailRests)
            {
                writer.WriteStartElement("field");
                WriteXmlValue(writer, "name", "retailRest");
                WriteXmlValue(writer, "description", "Остаток (розница)");
                WriteXmlValue(writer, "type", "string");
                WriteXmlValue(writer, "length", "30");
                WriteXmlValue(writer, "align", "1");
                writer.WriteEndElement();

                writer.WriteStartElement("field");
                WriteXmlValue(writer, "name", "recommendOrder");
                WriteXmlValue(writer, "description", "Заказ рекомендуемый");
                WriteXmlValue(writer, "type", "string");
                WriteXmlValue(writer, "length", "30");
                WriteXmlValue(writer, "align", "1");
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public static void WriteGroups(XmlWriter writer, List<ProductGroup> clientProductGroups)
        {
            if (clientProductGroups.Count == 0)
                return;

            writer.WriteStartElement("groups");

            foreach (ProductGroup group in clientProductGroups)
            {
                writer.WriteStartElement("group");
                WriteXmlValue(writer, "name", group.description);
                WriteXmlValue(writer, "level", group.sLevel);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public static void WriteData(XmlWriter writer, Client client, List<Product> clientProducts, bool useRetailRests, List<RetailRest>clientRetailRests)
        {
            writer.WriteStartElement("data");

            foreach (var product in clientProducts)
            {
                writer.WriteStartElement("row");
                WriteXmlValue(writer, "code", product.code);
                if (client.isExportProductArticle)
                    WriteXmlValue(writer, "n_article", product.article);
                WriteXmlValue(writer, "description", product.description);
                WriteXmlValue(writer, "unit", product.unit);
                WriteXmlValue(writer, "quantity", product.quantity);
                WriteXmlValue(writer, "pack_coefficient", product.pack);
                WriteXmlValue(writer, "price", product.price.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture));
                WriteXmlValue(writer, "currency", "643");
                if (client.isExportByProductMatrix)
                    WriteXmlValue(writer, "recommendQuantity", product.recommendQuantityString);
                WriteXmlValue(writer, "level", product.level);
                if (useRetailRests)
                {
                    RetailRest retailRest = clientRetailRests.Where(w => w.productCode == product.code).FirstOrDefault();

                    if ((retailRest?.rest ?? 0) == 0)
                        WriteXmlValue(writer, "retailRest", "--");
                    else
                        WriteXmlValue(writer, "retailRest", $"{retailRest.rest.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)} {retailRest.unit}");

                    if (retailRest == null)
                        WriteXmlValue(writer, "recommendOrder", $"{product.recommendQuantity.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)} {product.unit}");
                    else if (retailRest != null && retailRest.coefficient == 0)
                        WriteXmlValue(writer, "recommendOrder", "??");
                    else
                    {
                        decimal restInProductUnit = 0;
                        if (retailRest.coefficient > 0)
                            restInProductUnit = retailRest.rest * retailRest.coefficient;
                        else
                            restInProductUnit = retailRest.rest / -retailRest.coefficient;
                        decimal recommendOrder = product.recommendQuantity - restInProductUnit;
                        if (recommendOrder <= 0)
                            WriteXmlValue(writer, "recommendOrder", "--");
                        else
                            WriteXmlValue(writer, "recommendOrder", $"{recommendOrder.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)} {product.unit}");
                    }
                }
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        public static void WriteXmlValue(XmlWriter writer, string tag, string value)
        {
            writer.WriteStartElement(tag);
            writer.WriteValue(value);
            writer.WriteEndElement();
        }
    }
}
