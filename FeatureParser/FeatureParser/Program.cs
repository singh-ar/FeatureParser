using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FeatureParser
{
    public class Program
    {
        private HashSet<string> keywordSet;
        private HashSet<string> filePathSet;
        private StreamWriter logFile;

        private Dictionary<string, KeywordData> keywordUsageData;
        
        private class KeywordData
        {
            public int ParameterCount { get; set; }
            public int UseCount { get; set; }
            //public Dictionary<string, List<string>> parameterData { get; set; }

            public HashSet<Dictionary<string, List<string>>> ParameterDataSet { get; set; }

            public KeywordData()
            {
                //parameterData = new Dictionary<string, List<string>>();
                ParameterDataSet = new HashSet<Dictionary<string, List<string>>>();
            }

            public string GetObjectString()
            {
                string str = "";
                
                foreach (var data in this.ParameterDataSet)
                {
                    string temp = "";
                    foreach (var dElement in data)
                    {
                        str += dElement.Key;
                        foreach (var param in dElement.Value)
                        {
                            str += $" :: {param}";
                        }

                        str += "\n";
                    }
                    str += "\n";
                }

                return str;
            }

            public bool VerifyAndAdd(Dictionary<string, List<string>> paramData)
            {
                bool match = false;

                foreach (var data in this.ParameterDataSet)
                {
                    // compare objects
                    if (data.Count == paramData.Count)
                    {
                        foreach (var dElement in data)
                        {
                            try
                            {
                                var list1 = paramData[dElement.Key];
                                var list2 = dElement.Value;

                                if (list1.Count == list2.Count)
                                {
                                    for (int i = 0; i < dElement.Value.Count; i++)
                                        if (list1[i].Equals(list2[i]))
                                            match = true;
                                        else
                                            match = false;
                                }
                            }
                            catch (KeyNotFoundException e)
                            {
                                
                            }
                        }
                    }
                }

                if(!match)
                    this.ParameterDataSet.Add(paramData);

                return match;
            }
        }


        Program()
        {
            this.logFile = new StreamWriter("C:/Development/FeatureParser/FeatureParser/log.txt");
            this.keywordUsageData = new Dictionary<string, KeywordData>();
            this.filePathSet = new HashSet<string>();
        }

        public void Log(string message)
        {
            logFile.WriteLine(message);
            logFile.Flush();
        }


        #region Keyword Collection
        // read all the files in a folder recursively
        private HashSet<string> GetFileNamesOfDirectoryRecursively(string sDir)
        {

            // get present directory files
            foreach (var f in Directory.GetFiles(sDir))
                this.filePathSet.Add(f);
            

            //get sub directories
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                    GetFileNamesOfDirectoryRecursively(d);
            }
            catch (System.Exception excpt)
            {
                Log(excpt.Message);
            }
            
            return this.filePathSet;
        }


        // Read all keywords from the file
        private void ReadGivenKeywords(string filePath)
        {
            if (this.keywordSet != null)
            {
                if (File.Exists(filePath))
                {
                    // Read file using StreamReader. Reads file line by line  
                    using (StreamReader file = new StreamReader(filePath))
                    {
                        string ln;
                        while ((ln = file.ReadLine()) != null)
                        {
                            if (ln.Contains("[Given("))
                            {
                                string temp = TrimStartString(TrimEndString(ln.Trim(), "\")]"), "[Given(@\"");
                                this.keywordSet.Add("^" + temp + "$");
                            }
                            // can be used if givens change
                            else if (ln.Contains("[When("))
                            {
                                string temp = TrimStartString(TrimEndString(ln.Trim(), "\")]"), "[When(@\"");
                                this.keywordSet.Add("^" + temp + "$");
                            }
                        }
                        file.Close();
                    }
                }
                else
                    Log("File does not exist.");
            }
        }


        // collect the given binding keywords
        private void CollectGivenKeywords()
        {
            this.keywordSet = new HashSet<string>();

            foreach (string s in this.filePathSet)
                ReadGivenKeywords(s);

            Log("\nTotal keywords found " + this.keywordSet.Count);

            //foreach (string s in this.keywordSet)
            //    Log(s);
        }
        #endregion
        
        #region Data Collection
        // prepare keywordUsageData object
        public void PrepareKeywordUsageDataObject()
        {
            foreach (string s in this.keywordSet)
            {
                this.keywordUsageData.Add(s.Trim(), new KeywordData());
            }
        }

        public void SaveToKeywordUsageDataObject(string key, Dictionary<string, List<string>> keyWordParams)
        {
            //Log(key);
            this.keywordUsageData[key].UseCount++;
            this.keywordUsageData[key].VerifyAndAdd(keyWordParams);
            //if (this.keywordUsageData[key].VerifyAndAdd(keyWordParams)) 
                //Log($"Similar object already present for --- {key}");
        }
        #endregion

        #region Parser
        // parse feature files
        public void ParseFeatureFile(string filePath)
        {
            //filePath = "C:/Development/FeatureParser/Configuration.feature";
            
            if (File.Exists(filePath))
            {
                // Read file using StreamReader. Reads file line by line  
                using (StreamReader file = new StreamReader(filePath))
                {
                    string ln;

                    // data to be stored
                    string latestKeyword = "";
                    int latestParameterCount = 0;
                    Dictionary<string, List<string>> paramDictionary = null;
                    Dictionary<string, Dictionary<string, List<string>>> mainData = new Dictionary<string, Dictionary<string, List<string>>>();


                    // local data
                    bool isCollectingParameters = false;
                    int parameterLineCounter = 0;
                    List<string> paramNames = null;


                    while ((ln = file.ReadLine()) != null)
                    {
                        string keyword = FindMatchingKeyword(ln.Trim());

                        if (keyword != null) // a keyword is found
                        {
                            if (isCollectingParameters) // save the keyword data here
                                SaveToKeywordUsageDataObject(latestKeyword, paramDictionary);

                            // reset parameters
                            isCollectingParameters = true;
                            latestKeyword = keyword;
                            latestParameterCount = 0;

                            parameterLineCounter = 0;

                            paramDictionary = new Dictionary<string, List<string>>();
                            paramNames = new List<string>();
                        }
                        else if (isCollectingParameters) // when reading parameters
                        {
                            if (ln.Trim().StartsWith("|")) // when new line has paramaters
                            {
                                parameterLineCounter++;

                                if (parameterLineCounter == 1) // when reading headers
                                {
                                    var headers = ln.Trim().Split("|");

                                    latestParameterCount = headers.Length - 2;

                                    for (int i = 1; i < (headers.Length - 1); i++)
                                    {
                                        paramNames.Add(headers[i].Trim());
                                        try
                                        {
                                            paramDictionary.Add(headers[i].Trim(), new List<string>());  // save keys for paramater data
                                        }
                                        catch (ArgumentException e)
                                        {
                                            Log($"File {filePath} has repeated parameter value");
                                        }
                                    }
                                }
                                else
                                {
                                    // save parameter values
                                    var values = ln.Trim().Split("|");

                                    for (int i = 1; i < (values.Length - 1); i++)
                                        paramDictionary[paramNames[i - 1]].Add(values[i].Trim());
                                }
                            }
                            else // when new line has no parameters
                            {
                                SaveToKeywordUsageDataObject(latestKeyword, paramDictionary);

                                // reset parameters
                                isCollectingParameters = false;
                                latestKeyword = "";

                                parameterLineCounter = 0;
                                paramDictionary = null;
                                paramNames = null;
                            }
                        }
                    }

                    if (isCollectingParameters && (ln = file.ReadLine()) == null) // save the keyword data here
                        SaveToKeywordUsageDataObject(latestKeyword, paramDictionary);
                    file.Close();
                }
            }
        }

        public void ParseAllFeatureFiles()
        {
            foreach (var path in this.filePathSet)
            {
                if(path.Trim().EndsWith(".feature"))
                    ParseFeatureFile(path);
            }
        }
        
        #endregion
        
        // Helper functions
        #region Helpers
        private string FindMatchingKeyword(string codeLine)
        {
            if (!codeLine.StartsWith("|") 
                && !codeLine.StartsWith("#") 
                && !codeLine.StartsWith("Scenario:") 
                && !codeLine.StartsWith("@") 
                && !codeLine.StartsWith("Feature:")
                && !codeLine.StartsWith("#")
                && !codeLine.StartsWith("Then")
                && !(codeLine == ""))
            {
                foreach (string s in this.keywordSet)
                {
                    string valueToMatch = "";

                    if (codeLine.StartsWith("When"))
                    {
                        valueToMatch = TrimStartString(codeLine,"When ");
                    }

                    if (codeLine.StartsWith("Given"))
                    {
                        valueToMatch = TrimStartString(codeLine, "Given ");
                    }

                    if (codeLine.StartsWith("And"))
                    {
                        valueToMatch = TrimStartString(codeLine, "And ");
                    }

                    if (Regex.IsMatch(valueToMatch.Trim(), s.Trim()))
                        return s;
                }
            }

            return null;
        }

        public void LogKeywordDictionary(Dictionary<string, List<string>> dictionary)
        {
            foreach (var s in dictionary)
            {
                Log(s.Key);
                foreach (string a in s.Value)
                    Log(a);
            }
        }

        // String manipulation
        public static string TrimStartString(string target, string trimString)
        {
            if (string.IsNullOrEmpty(trimString)) return target;

            string result = target;
            while (result.StartsWith(trimString))
            {
                result = result.Substring(trimString.Length);
            }

            return result.Trim();
        }

        public static string TrimEndString(string target, string trimString)
        {
            if (string.IsNullOrEmpty(trimString)) return target;

            string result = target;
            while (result.EndsWith(trimString))
            {
                result = result.Substring(0, result.Length - trimString.Length);
            }

            return result.Trim();
        }

        public void LogKeywordUsage()
        {
            int count = 0;
            foreach (var a in this.keywordUsageData)
            {
                if (a.Value.UseCount > 0)
                {
                    Log($"{a.Key} --- {a.Value.UseCount}");
                    count++;
                }
            }
            Log($"{count} keywords used out of {keywordSet.Count}");
        }

        public void PrintAllFileName()
        {
            foreach (string s in this.filePathSet)
                Log(s);
        }

        private List<int> GetParamColumnNumList(Dictionary<string, int> paramColNumDictionary, Dictionary<string, List<string>> data)
        {
            List<int> returnList = new List<int>();

            if (paramColNumDictionary.Count == 0) // no param name saved yet
            {
                int colNum = 0;

                foreach (var d in data)
                {
                    returnList.Add(colNum);
                    paramColNumDictionary.Add(d.Key,colNum);
                    colNum++;
                }
            }
            else
            {
                int colNum = paramColNumDictionary.Count;
                foreach (var d in data)
                {
                    if (paramColNumDictionary.ContainsKey(d.Key))
                    {
                        returnList.Add(paramColNumDictionary[d.Key]);
                    }
                    else
                    {
                        returnList.Add(paramColNumDictionary.Count);
                        paramColNumDictionary.Add(d.Key,paramColNumDictionary.Count);
                    }
                }
            }

            return returnList;
        }

        public void WriteToExcel()
        {
            XLWorkbook workbook = new XLWorkbook();
            DataTable table = new DataTable();
            List<int> columnNumbers;

            // add column names
            for (int i = 0; i < 30; i++)
                table.Columns.Add(i.ToString());

            string keyword = "";
            int useCount = 0;

            foreach (var a in this.keywordUsageData)
            {
                Dictionary<string, int> paramColNumDictionary = new Dictionary<string, int>();
                if (a.Value.UseCount > 0)
                {
                    var keyRow = table.NewRow();
                    keyRow[0] = a.Key;
                    keyRow[1] = a.Value.UseCount;

                    table.Rows.Add(keyRow);

                    // loop through dataset
                    foreach (var data in a.Value.ParameterDataSet)
                    {
                        columnNumbers = GetParamColumnNumList(paramColNumDictionary, data);

                        if (data.Count > 0)
                        {
                            int rowCounter = data.First().Value.Count;

                            // loop through rows
                            for (var rowCount = 0; rowCount <= rowCounter; rowCount++)
                            {
                                var row = table.NewRow();
                                int i = 0;


                                foreach (var paramName in data) // add parameter names
                                {
                                    if (rowCount == 0)
                                    {
                                        row[columnNumbers[i] + 2] = paramName.Key;
                                    }
                                    else
                                    {
                                        //row[i + 2] = paramName.Value.ElementAt(rowCount - 1);
                                        row[columnNumbers[i] + 2] = paramName.Value.ElementAt(rowCount - 1);
                                    }
                                    i++;
                                }

                                table.Rows.Add(row);
                            }

                            table.Rows.Add(table.NewRow());
                        }
                    }
                }
            }

            workbook.Worksheets.Add(table, "name");
            workbook.SaveAs("C:/Development/FeatureParser/data.xlsx");
        }
        
        #endregion


        static async Task Main(string[] args)
        {
            Program obj = new Program();

            obj.GetFileNamesOfDirectoryRecursively("C:/Development/Promapp/Promapp.Test.UI/Steps/CommonSetupSteps");
            obj.CollectGivenKeywords();

            obj.PrepareKeywordUsageDataObject();

            //obj.ParseFeatureFile("");

            obj.GetFileNamesOfDirectoryRecursively("C:/Development/Promapp/Promapp.Test.UI/SpecFlow");
            //obj.PrintAllFileName();

            obj.ParseAllFeatureFiles();
            //obj.LogKeywordUsage();
            //obj.WriteToExcel();
            obj.WriteToExcel();

        }
    }
}
