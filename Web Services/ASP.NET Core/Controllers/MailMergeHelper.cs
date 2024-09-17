using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;

namespace EJ2APIServices
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
                if (items.ContainsKey("StartNextRecordAtNewPage"))
                    wordDocument.MailMerge.StartAtNewPage = bool.Parse(items["StartNextRecordAtNewPage"].ToString());
                else
                    wordDocument.MailMerge.StartAtNewPage = false;
                if (items.ContainsKey("MergeType"))
                {
                    switch (items["MergeType"].ToString())
                    {
                        case "Simple":
                            ExecuteSimpleMerge(wordDocument, items);
                            break;
                        case "Group":
                            ExecuteGroupMerge(wordDocument, items);
                            break;
                        case "NestedGroup":
                            ExecuteNestedGroupMerge(wordDocument, items);
                            break;
                    }
                }
            }
        }
        private static void ExecuteSimpleMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                string tableName = items["MergeCommand"].ToString();
                // Gets JSON object from JSON string.
                JObject jsonObject = JObject.Parse(items["DataSource"].ToString());
                // Converts to IDictionary data from JSON object.
                IDictionary<string, object> dataSource = GetData(jsonObject);
                if (dataSource is IDictionary<string, object>)
                {
                    //Executes Mail merge.
                    wordDocument.MailMerge.Execute(dataSource);
                }
            }
        }
        private static void ExecuteGroupMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                string tableName = items["MergeCommand"].ToString();
                // Gets JSON object from JSON string.
                JObject jsonObject = JObject.Parse(items["DataSource"].ToString());
                // Converts to IDictionary data from JSON object.
                IDictionary<string, object> dataSource = GetData(jsonObject);
                List<object> dataCollection;
                if (dataSource is IDictionary<string, object> && dataSource.ContainsKey(tableName))
                {
                    dataCollection = dataSource[tableName] as List<object>;
                    MailMergeDataTable dataTable = new MailMergeDataTable(tableName, dataCollection);
                    //Executes Mail merge for group.
                    wordDocument.MailMerge.ExecuteGroup(dataTable);
                }
            }
        }
        private static void ExecuteNestedGroupMerge(WordDocument wordDocument, Dictionary<string, object> items)
        {
            if (items.ContainsKey("MergeCommand"))
            {
                string tableName = items["MergeCommand"].ToString();
                // Gets JSON object from JSON string.
                JObject jsonObject = JObject.Parse(items["DataSource"].ToString());
                // Converts to IDictionary data from JSON object.
                IDictionary<string, object> dataSource = GetData(jsonObject);
                if (dataSource is IDictionary<string, object> && dataSource.ContainsKey(tableName))
                {
                    List<object> dataCollection = dataSource[tableName] as List<object>;
                    if (items.ContainsKey("RecordId"))
                    {
                        int recordId = int.Parse(items["RecordId"].ToString());
                        for (int i = 0; i < dataCollection.Count; i++)
                        {
                            if (recordId == i)
                            {
                                List<object> tempDataCollection = new List<object>();
                                tempDataCollection.Add(dataCollection[i]);
                                MailMergeDataTable dataTable = new MailMergeDataTable(tableName, tempDataCollection);
                                //Executes Mail merge for nested group.
                                wordDocument.MailMerge.ExecuteNestedGroup(dataTable);
                                return;
                            }
                        }
                    }
                    else
                    {
                        MailMergeDataTable dataTable = new MailMergeDataTable(tableName, dataCollection);
                        //Executes Mail merge for nested group.
                        wordDocument.MailMerge.ExecuteNestedGroup(dataTable);
                    }
                }
            }
        }
        /// <summary>
        /// Gets array of items from JSON array.
        /// </summary>
        /// <param name="jArray">JSON array.</param>
        /// <returns>List of objects.</returns>
        private static List<object> GetData(JArray jArray)
        {
            List<object> jArrayItems = new List<object>();
            foreach (var item in jArray)
            {
                object keyValue = null;
                if (item is JObject)
                    keyValue = GetData((JObject)item);
                jArrayItems.Add(keyValue);
            }
            return jArrayItems;
        }
        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>IDictionary data.</returns>
        private static IDictionary<string, object> GetData(JObject jsonObject)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            foreach (var item in jsonObject)
            {
                object keyValue = null;
                if (item.Value is JArray)
                    keyValue = GetData((JArray)item.Value);
                else if (item.Value is JToken)
                {
                    if (((JToken)item.Value).Type == JTokenType.Object)
                    {
                        keyValue = GetData((JObject)item.Value);
                        List<object> jArrayItems = new List<object>();
                        jArrayItems.Add(keyValue);
                        keyValue = jArrayItems;
                    }
                    else
                    {
                        keyValue = ((JToken)item.Value).ToObject<string>();
                    }
                }
                dictionary.Add(item.Key, keyValue);
            }
            return dictionary;
        }
    }
}
