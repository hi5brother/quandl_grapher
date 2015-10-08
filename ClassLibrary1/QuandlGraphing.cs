using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using System.Runtime.InteropServices;

namespace QuandlProcess
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("QuandlGrapher")]
    public class Grapher
    {
        public DataAccess data = new DataAccess();
        private Boolean initiated = false;
        public Boolean DataSet(string quandlDatabase, string datacodeParams, int datapointsNumber, string frequency)
        {

            data.InitiateDataset(quandlDatabase, datacodeParams, datapointsNumber, frequency);
            initiated = true;
            return true;

        }
        public void CreateSingleGraph()
        {
            //http://clear-lines.com/blog/post/Create-an-Excel-chart-in-C-without-worksheet-data.aspx
            Excel.Application xl;
            Excel.Workbook wb;
            Excel.Worksheet ws;


            xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            xl.Visible = true;
            wb = (Excel.Workbook)xl.ActiveWorkbook;
            ws = (Excel.Worksheet)wb.Sheets["summary"];

            ////Creating the chart
            Excel.ChartObjects nChart = (Excel.ChartObjects)ws.ChartObjects(Type.Missing);
            Excel.ChartObject chartObj = (Excel.ChartObject)nChart.Add(10, 30, 200, 200);
            Excel.Chart chart = chartObj.Chart;

            chart.ChartType = Excel.XlChartType.xlLine;

            var seriesCol = chart.SeriesCollection();
            var series = seriesCol.NewSeries();

            if (this.initiated)
            {
                float[] valArr = data.ReturnFloatValues("Close");
                Array.Reverse(valArr);
                series.Values = valArr;

                string[] datesArr = data.ReturnDates();
                Array.Reverse(datesArr);
                series.XValues = datesArr;

                series.Name = data.ReturnName();
            }


        }
    }
    public class TechnicalAnalysis
    {
        //public float[] ReturnMovingAverages(float[] values)
        //public float[] ReturnUpperBollingerBand(float[] values)
        //public float[] ReturnLowerBollingerBand(float[] values)
    }
    
}
