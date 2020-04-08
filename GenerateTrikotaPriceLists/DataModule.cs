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

        public class ProductGroup: ICloneable
        {
            public string code { get; set; }
            public string description { get; set; }
            public int iLevel { get; set; }
            public string sLevel { get; set; }

            public ProductGroup() { }

            public ProductGroup(string groups)
            {
                var values = groups.Split('|');
                code = values[0];
                description = values[1];
                iLevel = StrToIntDef(values[2], 1);
                sLevel = "";
            }

            public ProductGroup(DataRow row)
            {
                code = (string)row["code"];
                description = (string)row["description"];
                iLevel = (int)row["iLevel"];
                sLevel = (string)row["slevel"];
            }

            public object Clone()
            {
                return this.MemberwiseClone();
            }
        }

        public class ProductMatrixElement 
        {
            public string companyCode { get; set; }
            public string storehouseCode { get; set; }
            public string productCode { get; set; }
            public string recommendQuantity { get; set; }

            public ProductMatrixElement() { recommendQuantity = ""; }
            public ProductMatrixElement(string line)
            {
                try
                {
                    string[] values = line.Split(';');

                    companyCode = values[0];
                    storehouseCode = values[1];
                    productCode = values[2];
                    recommendQuantity = values[3];
                }
                catch (Exception exception)
                {
                    throw new Exception($"Некорректный файл с товарной матрицей:\n{exception.Message}");
                }
            }
        }

        public class Product : ICloneable
        {
            public static int fieldsCount = 8;

            public string code { get; set; }
            public string article { get; set; }
            public string description { get; set; }
            public string unit { get; set; }
            public string pack { get; set; }
            public string characteristicDescription { get; set; }
            public string quantity { get; set; }
            public string recommendQuantity { get; set; }
            public string level { get; set; }
            public decimal price { get; set; }
            public string brand { get; set; }
            public string barcode { get; set; }
            public string comment { get; set; }
            public List<ProductGroup> groups;

            public Product() { }

            public Product(string line)
            {
                try
                {
                    string[] values = line.Split(';');

                    int i = 0;
                    code = values[i++];
                    article = values[i++];
                    description = values[i++];
                    brand = values[i++];
                    unit = values[i++];
                    barcode = values[i++];
                    pack = values[i++];
                    characteristicDescription = values[i++];
                    quantity = values[i++];
                    recommendQuantity = "";

                    level = "";
                    price = 0;

                    groups = new List<ProductGroup>();
                    foreach (var group in values[i++].Split('#'))
                    {
                        groups.Add(new ProductGroup(group));
                    }

                    comment = values[i++].Replace("\\n", "\n"); ;
                }
                catch (Exception exception)
                {
                    throw new Exception($"Некорректный файл с остатками номенклатуры:\n{exception.Message}");
                }
            }

            public object Clone()
            {
                return new Product()
                {
                    code = this.code,
                    article = this.article,
                    description = this.description,
                    unit = this.unit,
                    pack = this.pack,
                    characteristicDescription = this.characteristicDescription,
                    quantity = this.quantity,
                    recommendQuantity = this.recommendQuantity,
                    level = this.level,
                    price = this.price,
                    comment = this.comment,
                    brand = this.brand,
                    barcode = this.barcode,
                    groups = this.groups.Select(s => (ProductGroup)s.Clone()).ToList<ProductGroup>()
                };
            }
        }

        public class ContractSpecialCondition
        {
            public string productCode { get; set; }
            public string characteristicDescription { get; set; }
            public string priceTypeCode { get; set; }
            public decimal discount { get; set; }

            public ContractSpecialCondition() { }
            public ContractSpecialCondition(string conditions)
            {
                var values = conditions.Split('|');
                productCode = values[0];
                characteristicDescription = values[1];
                priceTypeCode = values[2];
                discount = StrToDecimalDef(values[3], 0);
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
            public int exportToEXCELVariant { get; set; }

            public string exportPath { get; set; }

            public List<ContractSpecialCondition> specialConditions;

            public Client() {  }

            public Client(string line)
            {
                try
                {
                    string[] values = line.Split(';');

                    int i = 0;
                    clientCode = values[i++];
                    clientDescription = values[i++];
                    contractCode = values[i++];
                    contractDescription = values[i++];
                    contractPriceTypeCode = values[i++];
                    contractDiscount = StrToDecimalDef(values[i++], 0);
                    isExportByContract = StrToBoolDef(values[i++], false);
                    isExportBySpecialConditionsProducts = StrToBoolDef(values[i++], false);
                    isExportByProductMatrix = StrToBoolDef(values[i++], false);
                    exportPath = values[i++];
                    groupDepth = StrToIntDef(values[i++], 0);
                    isExportProductArticle = StrToBoolDef(values[i++], false);
                    isExportProductComment = StrToBoolDef(values[i++], false);
                    isAppendClientCodeExcel = StrToBoolDef(values[i++], false);
                    companyCodeForMatrixFilter = values[i++];
                    storehouseCodeForMatrixFilter = values[i++];
                    isExportToXML = StrToBoolDef(values[i++], false);
                    isExportToEXCEL = StrToBoolDef(values[i++], false);
                    exportToEXCELVariant = StrToIntDef(values[i++], 1);

                    specialConditions = new List<ContractSpecialCondition>();
                    foreach (var condition in values[i++].Split('#'))
                    {
                        if (String.IsNullOrWhiteSpace(condition))
                            continue;

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
        public static List<Product> allMatrixProducts;
        public static List<ProductGroup> productGroups = new List<ProductGroup>();
        public static List<ProductMatrixElement> productMatrix;
        public static Dictionary<string, Dictionary<string, decimal>> productPrices;
        public static List<Client> clients;

        public static void LoadConstants(string path)
        {
            logger.Trace("Загрузка констант...");

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

        public static void LoadMatrix(string path)
        {
            logger.Trace("Загрузка товарной матрицы...");

            productMatrix = new List<ProductMatrixElement>();

            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(1251));

                foreach (string line in lines)
                {
                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    productMatrix.Add(new ProductMatrixElement(line));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadAllMatrixProducts(string path)
        {
            logger.Trace("Загрузка номенклатуры матрицы...");

            allMatrixProducts = new List<Product>();

            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(1251));

                foreach (string line in lines)
                {
                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    allMatrixProducts.Add(new Product(line));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadProducts(string path)
        {
            logger.Trace("Загрузка остатков номенклатуры...");

            products = new List<Product>();

            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(1251));

                foreach (string line in lines)
                {
                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    products.Add(new Product(line));
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadProductPrices(string path)
        {
            logger.Trace("Загрузка цен номенклатуры...");

            productPrices = new Dictionary<string, Dictionary<string, decimal>>();

            try
            {
                string priceTypeCode, productCode;

                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(1251));

                foreach (string line in lines)
                {
                    string[] values = line.Split(';');

                    productCode = values[0];
                    priceTypeCode = values[1];

                    if (!productPrices.ContainsKey(productCode))
                        productPrices[productCode] = new Dictionary<string, decimal>();

                    productPrices[productCode][priceTypeCode] = StrToDecimalDef(values[2], 0);
                }
            }
            catch
            {
                throw;
            }
        }

        public static void LoadClients(string path)
        {
            logger.Trace("Загрузка списка контрагентов...");

            clients = new List<Client>();

            try
            {
                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(1251));

                foreach (string line in lines)
                {
                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    clients.Add(new Client(line));
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

        public static void FillProductGroups(List<Product> products, List<ProductGroup> productGroups)
        {
            logger.Trace("Подготовка данных...");

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
                logger.Trace($"Подготовка данных для {client.clientDescription}...");

                List<Product> clientProductsXml, clientProductsExcel;
                List<ProductGroup> clientProductGroupsXml, clientProductGroupsExcel;

                if (client.isExportToXML)
                {
                    clientProductsXml = products.Select(s => (Product)s.Clone()).ToList<Product>();
                    if (client.isExportByProductMatrix)
                    {
                        foreach (Product p in allMatrixProducts)
                        {
                            if (clientProductsXml.Find(f => f.code == p.code) == null)
                                clientProductsXml.Add((Product)p.Clone());
                        }
                        clientProductGroupsXml = new List<ProductGroup>();
                        FillProductGroups(clientProductsXml, clientProductGroupsXml);
                        clientProductGroupsXml = clientProductGroupsXml.Where(w => w.iLevel <= client.groupDepth).Select(s => (ProductGroup)s.Clone()).ToList<ProductGroup>();
                        clientProductsXml.All(p => { p.level = String.Join(".", p.level.Split('.').Take(client.groupDepth)); p.price = GetPrice(client, p); return true; });
                    } else {
                        if (client.groupDepth == 0)
                        {
                            clientProductGroupsXml = new List<ProductGroup>();
                            clientProductsXml.All(p => { p.level = "0"; p.price = GetPrice(client, p); return true; });
                        }
                        else
                        {
                            clientProductGroupsXml = productGroups.Where(w => w.iLevel <= client.groupDepth).Select(s => (ProductGroup)s.Clone()).ToList<ProductGroup>();
                            clientProductsXml.All(p => { p.level = String.Join(".", p.level.Split('.').Take(client.groupDepth)); p.price = GetPrice(client, p); return true; });
                        }
                    }

                    FilterProductsBySpecialConditions(client, ref clientProductsXml, ref clientProductGroupsXml);

                    ExportToXML.DoExportToXML(client, clientProductGroupsXml, clientProductsXml);
                }

                // Подразумевается, что при выгрузке в Excel не фильтруем по вложенности,
                // и отсутствует номенклатура, не входящая в какую-либо группу
                if (client.isExportToEXCEL)
                {
                    clientProductsExcel = products.Select(s => (Product)s.Clone()).ToList<Product>();
                    clientProductsExcel.All(p => { p.price = GetPrice(client, p); return true; });
                    clientProductGroupsExcel = productGroups.Select(s => (ProductGroup)s.Clone()).ToList<ProductGroup>();

                    FilterProductsBySpecialConditions(client, ref clientProductsExcel, ref clientProductGroupsExcel);

                    ExportToExcel.DoExportToExcel(client, clientProductsExcel, ExportToExcel.GetPreparedTable(clientProductGroupsExcel));
                }
            }
        }

        public static void FilterProductsBySpecialConditions(Client client, ref List<Product> clientProducts, ref List<ProductGroup> clientProductGroups)
        {
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
                clientProducts.All(p =>
                {
                    ProductMatrixElement matrixElement = productMatrix
                    .Where(w =>
                      w.companyCode == client.companyCodeForMatrixFilter &&
                      w.storehouseCode == client.storehouseCodeForMatrixFilter &&
                      w.productCode == p.code).FirstOrDefault();
                    p.recommendQuantity = matrixElement?.recommendQuantity ?? "";
                    return true; });
            }

            if (client.isExportBySpecialConditionsProducts || client.isExportByProductMatrix)
            {
                var groupCodes = clientProducts.SelectMany(s => s.groups).Select(s => s.code).Distinct();
                clientProductGroups = clientProductGroups.Where(g => groupCodes.Contains(g.code)).ToList<ProductGroup>();
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

            var specCondidionCharacteristic = client.specialConditions.Where(w => !String.IsNullOrWhiteSpace(w.characteristicDescription) && w.characteristicDescription == product.characteristicDescription).FirstOrDefault();
            if (specCondidionCharacteristic != null)
                return GetPrice(prices, specCondidionCharacteristic.priceTypeCode, specCondidionCharacteristic.discount);

            for (int i = product.groups.Count - 1; i >= 0; i--)
            {
                var specConditionProductGroup = client.specialConditions.Where(w => w.productCode == product.groups[i].code).FirstOrDefault();
                if (specConditionProductGroup != null)
                    return GetPrice(prices, specConditionProductGroup.priceTypeCode, specConditionProductGroup.discount);
            }

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
