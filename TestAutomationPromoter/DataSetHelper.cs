using System;
using System.Collections.Generic;
using System.IO;


namespace TestAutomationPromoter
{
    public static class DataSetHelper
    {
        public static void PromoteDataSet(string path, List<ExcelObject> ConfigObjectcollection)
        {
            string[] files = Directory.GetFiles(path, "*.xml", SearchOption.AllDirectories);

            foreach (var item in files)
            {
                FileWriter.WriteFiles(item, ConfigObjectcollection);
            }


        }
    }
}
