using Newtonsoft.Json;
using Syncfusion.DocIO.DLS;

namespace MailMergeExample
{
    public static class MailMergeHelper
    {
        public static void ExecuteMailMerge(this WordDocument wordDocument, string mergeData)
        {
            Dictionary<string, object> data = JsonConvert.DeserializeObject<Dictionary<string, object>>(mergeData);
            List<object> mergeSettings = JsonConvert.DeserializeObject<List<object>>(data["MergeSettings"].ToString());
            foreach (var mergeSetting in mergeSettings)
            {
                Dictionary<string, object> items = JsonConvert.DeserializeObject<Dictionary<string, object>>(mergeSetting.ToString());
                if (items.ContainsKey("ClearFields"))
                    wordDocument.MailMerge.ClearFields = bool.Parse(items["ClearFields"].ToString());
                else
                    wordDocument.MailMerge.ClearFields = true;
                if (items.ContainsKey("RemoveEmptyParagraphs"))
                    wordDocument.MailMerge.RemoveEmptyParagraphs = bool.Parse(items["RemoveEmptyParagraphs"].ToString());
                else
                    wordDocument.MailMerge.RemoveEmptyParagraphs = false;
                if (items.ContainsKey("RemoveEmptyGroup"))
                    wordDocument.MailMerge.RemoveEmptyGroup = bool.Parse(items["RemoveEmptyGroup"].ToString());
                else
                    wordDocument.MailMerge.RemoveEmptyGroup = false;
                if (items.ContainsKey("InsertAsNewRow"))
                    wordDocument.MailMerge.InsertAsNewRow = bool.Parse(items["InsertAsNewRow"].ToString());
                else
                    wordDocument.MailMerge.InsertAsNewRow = false;
                if (items.ContainsKey("StartAtNewPage"))
                    wordDocument.MailMerge.StartAtNewPage = bool.Parse(items["StartAtNewPage"].ToString());
                else
                    wordDocument.MailMerge.StartAtNewPage = false;
                if (items.ContainsKey("MergeType"))
                {
                    switch (items["MergeType"].ToString())
                    {
                        case "Simple":
                            break;
                        case "Group":
                            break;
                        case "NestedGroup":
                            break;
                    }
                    
                }
            }
        }
        private static void ExecuteSimpleMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                wordDocument.MailMerge.Execute(null);
            }
        }
        private static void ExecuteGroupMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                wordDocument.MailMerge.ExecuteGroup(null);
            }
        }
        private static void ExecuteNestedGroupMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                wordDocument.MailMerge.ExecuteNestedGroup(null);
            }
        }
    }
}
