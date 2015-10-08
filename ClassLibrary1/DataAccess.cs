using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.ComInterop;

namespace QuandlProcess
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("QuandlDataAccess")]
    //Interface with VBA as
    //Dim lib As Object: Set lib = CreateObject("QuandlProcess")
    //See: http://mikejuniperhill.blogspot.ca/2014/03/interfacing-c-and-vba-with-exceldna-no.html
    //
    public class DataAccess
    {
        private ParsedData quandlObject;

        public Boolean InitiateDataset(string quandlDatabase, string datacodeParams, int datapointsNumber, string frequency)
        {
            //create the object with the appropriate parameters
            this.quandlObject = QuandlProcess.GrabData(quandlDatabase, datacodeParams, datapointsNumber, frequency);
            return true;
        }
        public Boolean ClearDataset()
        {
            //delete the current object with the data
            this.quandlObject.Dispose();
            return true;
        }
        public String ReturnName()
        {
            return this.quandlObject.ReturnName();
        }

        public String GetDatatype(string key)
        {
            return this.quandlObject.ReturnDataType(key);
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public float[] ReturnFloatValues(String param)
        {
            return this.quandlObject.ReturnFloatValues(param);
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public String ReturnParamType(String param)
        {
            return this.quandlObject.ReturnDataType(param);
        }

        [return: MarshalAs(UnmanagedType.IDispatch)]
        public String[] ReturnDates()
        {
            return this.quandlObject.ReturnDates();
        }

        public float FindMax(String param)
        {
            //find max close point in the past ___ time period
            return this.quandlObject.ReturnFloatValues(param).Max();
        }
        public void FindMin(double timePeriod)
        {
            //find min close point in the past ___ time period
        }

    }

}
