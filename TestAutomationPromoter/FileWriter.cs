using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;


namespace TestAutomationPromoter
{
    public static class FileWriter
    {

        public static void WriteFiles(string filepath, List<ExcelObject> ConfigObjectcollection)
        {

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filepath);
            foreach (var config in ConfigObjectcollection)
            {

                var result = Path.GetFileName(filepath).ToLower();
                if (config.FileNameFilter != null)
                {
                    if (!result.Contains(config.FileNameFilter.ToLower())) continue;
                }

                var rootNode = "//" + config.EntityName;
                XmlNode rootCollection = xmlDoc.SelectSingleNode(rootNode);
                if (rootCollection == null) continue;
                if (xmlDoc.SelectSingleNode("//" + config.Tag) != null) continue;
                XmlNode xmlRecordNo = xmlDoc.CreateNode(XmlNodeType.Element, config.Tag, null);
                xmlRecordNo.InnerText = config.Value;
                rootCollection.InsertAfter(xmlRecordNo, rootCollection.LastChild);
                xmlDoc.Save(filepath);
                Console.WriteLine(result + " is modified to add " + config.Tag);
            }

        }

    }
}
