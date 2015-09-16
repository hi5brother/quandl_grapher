using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using QuandlCS.Requests;
using QuandlCS.Types;
using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace QuandlProcess
{

    public class QuandlProcess
    {
        public static string FindMax()
        {
            //string quandlDatabase = "PRAGUESE";
            //string datacodeParams = "PX";
            //int datapointsNumber = 150;
            ParsedData test = new ParsedData();
            test = GrabData("PRAGUESE", "PX", 150);
            return test.ReturnColumn("P")[0];

        }
        [ExcelFunction(Description="Grab Quandl Data into Lists")]
        public static ParsedData GrabData(string quandlDatabase, string datacodeParams, int datapointsNumber)
        {
            ////Hardcoded parameters
            //string quandlDatabase = "PRAGUESE";
            //string datacodeParams = "PX";
            //int datapointsNumber = 150;
            
            //Quandl request
            QuandlDownloadRequest request = new QuandlDownloadRequest();
            request.APIKey = "xNA_rA8KzZepxFUeu9bA";

            request.Datacode = new Datacode(quandlDatabase, datacodeParams); // PRAGUESE is the source, PX is the datacode
            request.Format = FileFormats.JSON;
            request.Frequency = Frequencies.Monthly;
            request.Truncation = datapointsNumber;
            request.Sort = SortOrders.Ascending;
            request.Transformation = Transformations.Difference;

            //OUTPUT: https://www.quandl.com/api/v1/datasets/PRAGUESE/PX.json?auth_token=xNA_rA8KzZepxFUeu9bA&collapse=monthly&transformation=diff&sort_order=asc&rows=150

            //Initialize data structure
            ParsedData pData = new ParsedData();
            List<string> paramList = new List<string>();
            List<string> paramType = new List<string>();

            //Add data parameter dictionary
            pData.Add("Date", "Heading");
            paramList.Add("Date");
            paramType.Add("datetype");

            //Add parameter dictionaries
            foreach (char parameter in datacodeParams)
            {
                pData.Add(parameter.ToString(), "Heading");
                paramList.Add(parameter.ToString());
            }
            

            //Parsing the data
            using (WebClient web = new WebClient())
            {
                string data = web.DownloadString(string.Format(request.ToRequestString()));
                JObject o = JObject.Parse(data);

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
                        
                        Debug.WriteLine(val.GetType());
                        //var convertedVal = System.Convert.ChangeType(val.ToString(), val.ToString().GetType());

                        pData.Add(paramList[count], val.ToString());

                        count++;
                    }
                }

                return pData;
                //C:\Users\Daniel\Documents\Visual Studio 2013\Projects\ClassLibrary1\ClassLibrary1\bin\Debug

           } 

        }
        //public static void WriteRange(int rows, int columns, Worksheet worksheet)
        //{
        //    for (var row = 1; row <= rows; rows ++)
        //    {
        //        for var column = 1; column <= column; column++)
        //        {
        //            var cell = (Range)worksheet.Cells[row, column];
        //            cell.Value2 = "Jokes";
        //        }
        //    }
        //}
        public static double Test(object[,] test)
        {
            return 2;
        }
    }
    public class ParsedData
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
        public void AddDataTypeList(List<string> dataTypeList)
        {
            int count = 0;
            foreach (KeyValuePair<string, List<string>> column in internalDictionary)
            {
                dataType.Add(column.Key, dataTypeList[count]);
                count++;
            }

        }


    }
}
