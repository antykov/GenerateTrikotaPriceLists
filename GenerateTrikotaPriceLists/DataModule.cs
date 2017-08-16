using NLog;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace GenerateTrikotaPriceLists
{
    public static class DataModule
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public class ProductGroup
        {
            public string code { get; set; }
            public string description { get; set; }
            public int iLevel { get; set; }
            public string sLevel { get; set; }

            public ProductGroup() { }

            public ProductGroup(string separatedValues)
            {
                var lines = separatedValues.Split('|');
                code = lines[0];
                description = lines[1];
                iLevel = StrToIntDef(lines[2], 1);
                sLevel = "";
            }

            public ProductGroup(DataRow row)
            {
                code = (string)row["code"];
                description = (string)row["description"];
                iLevel = (int)row["iLevel"];
                sLevel = (string)row["slevel"];
            }
        }

        public class ProductMatrixElement 
        {
            public string companyCode { get; set; }
            public string storehouseCode { get; set; }
            public string productCode { get; set; }

            public ProductMatrixElement() { }
            public ProductMatrixElement(string indexedTableValue)
            {
                try
                {
                    Regex regex = new Regex(@"(S:)(.*)");
                    MatchCollection matches = regex.Matches(indexedTableValue);

                    companyCode = matches[0].Groups[2].Value.Trim();
                    storehouseCode = matches[1].Groups[2].Value.Trim();
                    productCode= matches[2].Groups[2].Value.Trim();
                }
                catch (Exception exception)
                {
                    throw new Exception($"Некорректный файл с товарной матрицей:\n{exception.Message}");
                }
            }
        }

        public class Product
        {
            public static int fieldsCount = 8;

            public string code { get; set; }
            public string article { get; set; }
            public string description { get; set; }
            public string unit { get; set; }
            public string pack { get; set; }
            public string characteristicDescription { get; set; }
            public string quantity { get; set; }
            public string level { get; set; }
            public decimal price { get; set; }
            public List<ProductGroup> groups;

            public Product() { }

            public Product(string indexedTableValue)
            {
                try
                {
                    Regex regex = new Regex(@"(S:)(.*)");
                    MatchCollection matches = regex.Matches(indexedTableValue);

                    code = matches[0].Groups[2].Value.Trim();
                    article = matches[1].Groups[2].Value.Trim();
                    description = matches[2].Groups[2].Value.Trim();
                    unit = matches[3].Groups[2].Value.Trim();
                    pack = matches[4].Groups[2].Value.Trim();
                    characteristicDescription = matches[5].Groups[2].Value.Trim();
                    quantity = matches[6].Groups[2].Value.Trim();

                    level = "";
                    price = 0;

                    groups = new List<ProductGroup>();
                    foreach (var group in matches[7].Groups[2].Value.Trim().Split(';'))
                    {
                        groups.Add(new ProductGroup(group));
                    }
                }
                catch (Exception exception)
                {
                    throw new Exception($"Некорректный файл с остатками номенклатуры:\n{exception.Message}");
                }
            }
        }

        public class ContractSpecialCondition
        {
            public string productCode { get; set; }
            public string characteristicDescription { get; set; }
            public string priceTypeCode { get; set; }
            public decimal discount { get; set; }

            public ContractSpecialCondition() { }
            public ContractSpecialCondition(string condition)
            {
                var lines = condition.Split('|');
                productCode = lines[0];
                characteristicDescription = lines[1];
                priceTypeCode = lines[2];
                discount = StrToDecimalDef(lines[3], 0);
            }
        }

        public class Client
        {
            public string clientCode { get; set; }
            public string clientDescription { get; set; }
            public string contractCode { get; set; }
            public string contractDescription { get; set; }

            public string contractPriceTypeCode { get; set; }
            public decimal contractDiscount { get; set; }

            public bool isExportProductArticle { get; set; }
            public bool isExportProductComment { get; set; }
            public bool isAppendClientCodeExcel { get; set; }

            public bool isExportByContract { get; set; }
            public bool isExportBySpecialConditionsProducts { get; set; }
            public bool isExportByProductMatrix { get; set; }
            
            public int groupDepth { get; set; }

            public string companyCodeForMatrixFilter { get; set; }
            public string storehouseCodeForMatrixFilter { get; set; }

            public bool isExportToXML { get; set; }
            public bool isExportToEXCEL { get; set; }

            public string exportPath { get; set; }

            public List<ContractSpecialCondition> specialConditions;

            public Client() {  }

            public Client(string indexedTableValue)
            {
                try
                {
                    Regex regex = new Regex(@"(S:)(.*)");
                    MatchCollection matches = regex.Matches(indexedTableValue);

                    int index = 0;

                    clientCode = matches[index++].Groups[2].Value.Trim();
                    clientDescription = matches[index++].Groups[2].Value.Trim();
                    contractCode = matches[index++].Groups[2].Value.Trim();
                    contractDescription = matches[index++].Groups[2].Value.Trim();
                    contractPriceTypeCode = matches[index++].Groups[2].Value.Trim();
                    contractDiscount = StrToDecimalDef(matches[index++].Groups[2].Value.Trim(), 0);
                    isExportByContract = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    isExportBySpecialConditionsProducts = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    isExportByProductMatrix = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    exportPath = matches[index++].Groups[2].Value.Trim();
                    groupDepth = StrToIntDef(matches[index++].Groups[2].Value.Trim(), 0);
                    isExportProductArticle = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    isExportProductComment = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    isAppendClientCodeExcel = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    companyCodeForMatrixFilter = matches[index++].Groups[2].Value.Trim();
                    storehouseCodeForMatrixFilter = matches[index++].Groups[2].Value.Trim();
                    isExportToXML = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);
                    isExportToEXCEL = StrToBoolDef(matches[index++].Groups[2].Value.Trim(), false);

                    specialConditions = new List<ContractSpecialCondition>();
                    foreach (var condition in matches[index++].Groups[2].Value.Trim().Split(';'))
                    {
                        specialConditions.Add(new ContractSpecialCondition(condition));
                    }
                }
                catch (Exception exception)
                {
                    throw new Exception($"Некорректный файл с настройками контрагентов:\n{exception.Message}");
                }
            }

        }

        public static Dictionary<string, string> constants;
        public static List<Product> products;
        public static List<ProductGroup> productGroups = new List<ProductGroup>();
        public static List<ProductMatrixElement> productMatrix;
        public static Dictionary<string, Dictionary<string, decimal>> productPrices;
        public static List<Client> clients;

        public static MatchCollection GetIndexedTableStructure(string path)
        {
            string text = File.ReadAllText(path, Encoding.GetEncoding(1251));

            string indexedTableString = "{IndexedTable:";
            if (!text.Substring(0, indexedTableString.Length).Equals(indexedTableString))
                throw new Exception("Некорректный файл с остатками номенклатуры!");

            Regex regex = new Regex(@"\{((.|\n)*?)\}");

            return regex.Matches(text, indexedTableString.Length - 1);
        }

        public static void LoadConstants(string path)
        {
            logger.Info("Загрузка констант...");

            constants = new Dictionary<string, string>();

            string text = File.ReadAllText(path, Encoding.GetEncoding(1251));

            foreach (var line in text.Split('\n'))
            {
                var parts = line.Split('=');
                constants[parts[0]] = parts[1].Trim().Replace("\\n", "\n");
            }
        }

        public static string GetConstant(string key)
        {
            if (constants.ContainsKey(key))
                return constants[key];
            else
                return "";
        }

        public static void LoadProductMatrix(string path)
        {
            logger.Info("Загрузка товарной матрицы...");

            productMatrix = new List<ProductMatrixElement>();

            try
            {
                foreach (Match matches in GetIndexedTableStructure(path))
                {
                    productMatrix.Add(new ProductMatrixElement(matches.Value));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadProducts(string path)
        {
            logger.Info("Загрузка остатков номенклатуры...");

            products = new List<Product>();

            try
            {
                foreach (Match matches in GetIndexedTableStructure(path))
                {
                    products.Add(new Product(matches.Value));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadProductPrices(string path)
        {
            logger.Info("Загрузка цен номенклатуры...");

            productPrices = new Dictionary<string, Dictionary<string, decimal>>();

            try
            {
                string priceTypeCode, productCode;
                foreach (Match matches in GetIndexedTableStructure(path))
                {
                    Regex regex = new Regex(@"(S:)(.*)");
                    MatchCollection elementsMatches = regex.Matches(matches.Value);

                    productCode = elementsMatches[0].Groups[2].Value.Trim();
                    priceTypeCode = elementsMatches[1].Groups[2].Value.Trim();

                    if (!productPrices.ContainsKey(productCode))
                        productPrices[productCode] = new Dictionary<string, decimal>();

                    productPrices[productCode][priceTypeCode] = StrToDecimalDef(elementsMatches[2].Groups[2].Value.Trim(), 0);
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadClients(string path)
        {
            logger.Info("Загрузка списка контрагентов...");

            clients = new List<Client>();

            try
            {
                foreach (Match matches in GetIndexedTableStructure(path))
                {
                    clients.Add(new Client(matches.Value));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void RecursiveFillPriceListGroupsTable(DataTable table, List<ProductGroup> groups)
        {
            if (groups.Count == 0)
                return;

            DataRow row = table.Select($"code = '{groups[0].code}'").FirstOrDefault();
            if (row == null)
            {
                DataTable childrenTable = new DataTable();
                childrenTable.Columns.Add("code");
                childrenTable.Columns.Add("description");
                childrenTable.Columns.Add("iLevel", Type.GetType("System.Int32"));
                childrenTable.Columns.Add("sLevel");
                childrenTable.Columns.Add("children", childrenTable.GetType());

                row = table.NewRow();
                row["code"] = groups[0].code;
                row["description"] = groups[0].description;
                row["iLevel"] = groups[0].iLevel;
                row["sLevel"] = "";
                row["children"] = childrenTable;

                table.Rows.Add(row);
            }

            RecursiveFillPriceListGroupsTable((DataTable)row["children"], groups.Skip(1).ToList<ProductGroup>());
        }

        public static void ConvertPriceListGroupsTableToList(ref List<ProductGroup> list, DataTable table, string level = "")
        {
            if (table.Rows.Count == 0)
                return;

            table.DefaultView.Sort = "description";
            table = table.DefaultView.ToTable();

            int i = 1;
            string empty = "", dot = ".";
            foreach (DataRow row in table.Rows)
            {
                row["sLevel"] = $"{level}{(level.Length == 0 ? empty : dot)}{i}";
                list.Add(new ProductGroup(row));

                ConvertPriceListGroupsTableToList(ref list, (DataTable)row["children"], (string)row["sLevel"]);

                i++;
            }
        }

        public static void FillProductGroups()
        {
            logger.Info("Подготовка данных...");

            DataTable table = new DataTable();
            table.Columns.Add("code");
            table.Columns.Add("description");
            table.Columns.Add("iLevel", Type.GetType("System.Int32"));
            table.Columns.Add("sLevel");
            table.Columns.Add("children", table.GetType());

            foreach (Product item in products)
                RecursiveFillPriceListGroupsTable(table, item.groups);

            ConvertPriceListGroupsTableToList(ref productGroups, table);

            foreach (Product product in products)
            {
                if (product.groups.Count == 0)
                    continue;

                var group = productGroups.Where(w => w.code == product.groups.Last().code).FirstOrDefault();
                if (group == null)
                    continue;

                product.level = group.sLevel;
            }
        }

        public static void ExportPriceLists()
        {
            foreach (Client client in clients)
            {
                logger.Info($"Подготовка данных для {client.clientDescription}...");

                List<Product> clientProducts = products.Select(s => s).ToList<Product>();
                List<ProductGroup> clientProductGroups;

                if (client.groupDepth == 0)
                {
                    clientProductGroups = new List<ProductGroup>();
                    clientProducts.All(p => { p.level = "0"; p.price = GetPrice(client, p); return true; });
                } else
                {
                    clientProductGroups = productGroups.Where(w => w.iLevel <= client.groupDepth).ToList<ProductGroup>();
                    clientProducts.All(p => { p.level = String.Join(".", p.level.Split('.').Take(client.groupDepth)); p.price = GetPrice(client, p); return true; });
                }

                if (client.isExportBySpecialConditionsProducts)
                    clientProducts = clientProducts.Where(p => CheckProductForSpecialConditions(client, p)).ToList<Product>();

                if (client.isExportByProductMatrix)
                {
                    List<ProductMatrixElement> clientMatrix = productMatrix
                        .Where(w => 
                            w.companyCode == client.companyCodeForMatrixFilter &&
                            w.storehouseCode == client.storehouseCodeForMatrixFilter)
                         .ToList<ProductMatrixElement>();
                    clientProducts = clientProducts.Where(p => clientMatrix.Where(w => w.productCode == p.code).Count() > 0).ToList<Product>();
                }

                if (client.isExportBySpecialConditionsProducts || client.isExportByProductMatrix)
                { 
                    var groupCodes = clientProducts.SelectMany(s => s.groups).Select(s => s.code).Distinct();
                    clientProductGroups = clientProductGroups.Where(g => groupCodes.Contains(g.code)).ToList<ProductGroup>();
                }
                
                if (client.isExportToXML)
                    ExportToXML.DoExportToXML(client, clientProductGroups, clientProducts);

                if (client.isExportToEXCEL)
                    ExportToExcel.DoExportToExcel(client, clientProducts, ExportToExcel.GetPreparedTable(clientProductGroups));
            }
        }

        public static bool CheckProductForSpecialConditions(Client client, Product product)
        {
            if (client.specialConditions.Where(w => w.productCode == product.code).Count() > 0)
                return true;

            for (int i = product.groups.Count - 1; i >= 0; i--)
                if (client.specialConditions.Where(w => w.productCode == product.groups[i].code).Count() > 0)
                    return true;

            if (client.specialConditions.Where(w => !String.IsNullOrWhiteSpace(w.characteristicDescription) && w.characteristicDescription == product.characteristicDescription).Count() > 0)
                return true;

            return false;
        }

        public static decimal GetPrice(Dictionary<string, decimal> prices, string priceTypeCode, decimal discount)
        {
            if (prices.ContainsKey(priceTypeCode))
                return (decimal)Math.Round((double)prices[priceTypeCode] * (100 - (double)discount) / 100.0, 2, MidpointRounding.AwayFromZero);
            else
                return 0;
        }

        public static decimal GetPrice(Client client, Product product)
        {
            var prices = productPrices.ContainsKey(product.code) ? productPrices[product.code] : null;
            if (prices == null)
                return 0;

            var specConditionProduct = client.specialConditions.Where(w => w.productCode == product.code).FirstOrDefault();
            if (specConditionProduct != null)
                return GetPrice(prices, specConditionProduct.priceTypeCode, specConditionProduct.discount);

            for (int i = product.groups.Count - 1; i >= 0; i--)
            {
                var specConditionProductGroup = client.specialConditions.Where(w => w.productCode == product.groups[i].code).FirstOrDefault();
                if (specConditionProductGroup != null)
                    return GetPrice(prices, specConditionProductGroup.priceTypeCode, specConditionProductGroup.discount);
            }

            var specCondidionCharacteristic = client.specialConditions.Where(w => !String.IsNullOrWhiteSpace(w.characteristicDescription) && w.characteristicDescription == product.characteristicDescription).FirstOrDefault();
            if (specCondidionCharacteristic != null)
                return GetPrice(prices, specCondidionCharacteristic.priceTypeCode, specCondidionCharacteristic.discount);

            return GetPrice(prices, client.contractPriceTypeCode, client.contractDiscount); 
        }

        public static int StrToIntDef(string s, int def)
        {
            int result;
            if (Int32.TryParse(s, out result))
                return result;
            else
                return def;
        }

        public static decimal StrToDecimalDef(string s, decimal def)
        {
            decimal result;
            if (Decimal.TryParse(s.Replace(",", "."), System.Globalization.NumberStyles.Currency, System.Globalization.CultureInfo.InvariantCulture, out result))
                return result;
            else
                return def;
        }

        public static bool StrToBoolDef(string s, bool def)
        {
            bool result;
            if (s.Equals("1"))
                return true;
            if (bool.TryParse(s, out result))
                return result;
            else
                return def;
        }

        public static string RemovePathInvalidChars(string path, string replaceString = "")
        {
            string result = path;

            char[] invalidChars = Path.GetInvalidPathChars();
            foreach (var c in invalidChars)
            {
                result = result.Replace(c.ToString(), replaceString);
            }

            return result;
        }
    }
}
