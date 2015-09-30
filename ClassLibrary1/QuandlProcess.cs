using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Net;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using QuandlCS.Requests;
using QuandlCS.Types;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace QuandlProcess
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("QuandlProcess")]
    //Interface with VBA as
        //Dim lib As Object: Set lib = CreateObject("QuandlProcess")
        //See: http://mikejuniperhill.blogspot.ca/2014/03/interfacing-c-and-vba-with-exceldna-no.html
        //
    public class QuandlProcess
    {
        public static ParsedData GrabData(string quandlDatabase, string datacodeParams, int datapointsNumber, string frequency)
        {
            //Quandl request
            QuandlDownloadRequest request = new QuandlDownloadRequest();
            request.APIKey = "xNA_rA8KzZepxFUeu9bA";

            request.Datacode = new Datacode(quandlDatabase, datacodeParams); // PRAGUESE is the source, PX is the datacode
            request.Format = FileFormats.JSON;

            switch (frequency)
            {
                case "d":
                    request.Frequency = Frequencies.Daily;
                    break;
                case "w":
                    request.Frequency = Frequencies.Weekly;
                    break;
                case "m":
                    request.Frequency = Frequencies.Monthly;
                    break;
                case "q":
                    request.Frequency = Frequencies.Quarterly;
                    break;
                case "a":
                    request.Frequency = Frequencies.Annualy;
                    break;
                default:
                    request.Frequency = Frequencies.Daily;
                    break;
            }
            
            request.Truncation = datapointsNumber;
            request.Sort = SortOrders.Descending;
            request.Transformation = Transformations.None;

            //OUTPUT: https://www.quandl.com/api/v1/datasets/PRAGUESE/PX.json?auth_token=xNA_rA8KzZepxFUeu9bA&collapse=monthly&transformation=diff&sort_order=asc&rows=150

            //Initialize data structure
            ParsedData pData = new ParsedData();
            List<string> paramList = new List<string>();
            List<string> paramType = new List<string>();

            //Parsing the data
            using (WebClient web = new WebClient())
            {
                string data = web.DownloadString(string.Format(request.ToRequestString()));
                JObject o = JObject.Parse(data);

                //Parse column_names
                foreach (string parameter in o["column_names"].Children())
                {
                    pData.Add(parameter.ToString(), "Heading");
                    paramList.Add(parameter.ToString());
                }

                //var headings = o["column_names"].Children();
                var results = o["data"].Children();

                //REWRITE THIS---------------------------
                //Find the parameters of the data
                foreach (var val in results)
                {
                    foreach (var type in val)
                    {
                        paramType.Add(type.Type.ToString());
                    }
                    if (paramType.Count > paramList.Count)
                        break;
                }
                pData.AddDataTypeList(paramType);
                //-----------------------------------------
                foreach (var dataPoint in results)
                {
                    int count = 0;
                    foreach (var val in dataPoint)
                    {
                        pData.Add(paramList[count], val.ToString());
                        count++;
                    }
                }
                return pData;
           } 

        }
    }
    public class ParsedData: IDisposable
    {

        private Dictionary<string, List<string>> internalDictionary = new Dictionary<string, List<string>>();
        private Dictionary<string, string> dataType = new Dictionary<string, string>();

        
        public void Add(string key, string value)
        {
            if (this.internalDictionary.ContainsKey(key))
            {
                List<string> list = this.internalDictionary[key];
                list.Add(value);

            }
            else
            {
                List<string> list = new List<string>();
                list.Add(value);
                this.internalDictionary.Add(key, list);
            }
        }
        public List<string> ReturnColumn(string key)
        {
            List<string> list = this.internalDictionary[key];
            return list;

        }
        public float[] ReturnFloatValues(String param)
        {

            List<string> quandlList = this.ReturnColumn(param);
            string paramType = this.ReturnDataType(param);
            int dataCount = this.ReturnDataCount();

            if (paramType == "Float")
            {
                float[] returnResult = new float[dataCount];
                for (int i = 1; i <= dataCount; i++)
                {
                    returnResult[i - 1] = (float)Convert.ToDouble(quandlList[i]);
                }
                return returnResult;
            }
            else
            {
                return new float[1] { 0 };
            }
        }
        public string[] ReturnDates()
        {
            List<string> quandlList = this.ReturnColumn("Date");
            int dataCount = this.ReturnDataCount();
            string[] datesList = new string[dataCount];
            for (int i = 1; i<= dataCount; i++)
            {
                datesList[i - 1] = quandlList[i];
            }
            return datesList;
        }
        public int ReturnDataCount()
        {
           int dataCount = this.internalDictionary["Date"].Count - 1;
           return dataCount;
        }
        public void AddDataTypeList(List<string> dataTypeList)
        {
            int count = 0;
            foreach (KeyValuePair<string, List<string>> column in internalDictionary)
            {
                dataType.Add(column.Key, dataTypeList[count]);
                count++;
            }
        }
        public List<string> ReturnColumnNames()
        {
            List<string> keysList = new List<string>();
            foreach (string key in internalDictionary.Keys)
            {
                keysList.Add(key);
            }
            return keysList;
        }

        public string ReturnDataType(string key)
        {
            return dataType[key];
        }
        public void Dispose()
        {
            Dispose();
        }
        
    }
    [ComVisible(false)]
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }
}
