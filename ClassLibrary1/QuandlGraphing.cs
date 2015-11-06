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
            

            if (this.initiated)
            {
                //Data points for returns
                //we splice each array to show only visible values
                //accounts for the gap left from doing moving averages
                var series = seriesCol.NewSeries();
                
                float[] valArr = data.ReturnFloatValues("Close");
                Array.Reverse(valArr);
                float[] visibleValues = new List<float>(valArr).GetRange(12, 88).ToArray();
                series.Values = visibleValues;

                string[] datesArr = data.ReturnDates();
                Array.Reverse(datesArr);
                string[] visibleDates = new List<string>(datesArr).GetRange(12, 88).ToArray();
                series.XValues = visibleDates;

                series.Name = data.ReturnName();

                //moving averages
                var movingAvgSeries = seriesCol.NewSeries();
                float[] movingAveragesArr = Analysis.ReturnMovingAverages(valArr, 12);

                movingAvgSeries.Values = movingAveragesArr;

                movingAvgSeries.XValues = visibleDates;

                //bollinger bands
                var highBBSeries = seriesCol.NewSeries();
                float[] highBBArr = Analysis.ReturnUpperBollingerBand(valArr, 12);

                highBBSeries.Values = highBBArr;

                highBBSeries.XValues = visibleDates;

                var lowBBSeries = seriesCol.NewSeries();
                float[] lowBBArr = Analysis.ReturnLowerBollingerBand(valArr, 12);

                lowBBSeries.Values = lowBBArr;

                lowBBSeries.XValues = visibleDates;

            }

        }

    }
    public static class Analysis
    {
        public static float[] ReturnMovingAverages(float[] values, int days)
        {
            //days specify the number of days used in moving average data point

            int i,j;
            float[] movingAvgArr = new float[values.Length - days];
            float sum = 0;

            //Go through the first couple of days
            //No values for moving average!
            for (i = 0; i < days; i++ )
            {
                sum += values[i];
            }
            //Calculate for moving averages
            for (i = 0; i < values.Length - days; i++)
            {
                //subtract out the oldest value used in moving average
                sum = sum - values[i];
                //add in the new value used in moving average
                sum = sum + values[i + days];

                movingAvgArr[i] = sum / days;
            }
            return movingAvgArr;
        }
        public static float[] ReturnUpperBollingerBand(float[] returnValues, int days)
        {
            int i;
            float[] movingAvgArr = ReturnMovingAverages(returnValues, days);
            float[] stDevArr = ReturnStandardDeviationArray(returnValues, days);
            float[] BollingerBandArr = new float[movingAvgArr.Length];

            for (i = 0; i<movingAvgArr.Length; i++)
            {
                BollingerBandArr[i] = movingAvgArr[i] + 2 * stDevArr[i];
            }
            return BollingerBandArr;
        }
        public static float[] ReturnLowerBollingerBand(float[] returnValues, int days)
        {
            int i;
            float[] movingAvgArr = ReturnMovingAverages(returnValues, days);
            float[] stDevArr = ReturnStandardDeviationArray(returnValues, days);
            float[] BollingerBandArr = new float[movingAvgArr.Length];

            for (i = 0; i<movingAvgArr.Length; i++)
            {
                BollingerBandArr[i] = movingAvgArr[i] - 2 * stDevArr[i];
            }
            return BollingerBandArr;
        
        }
        private static float[] ReturnStandardDeviationArray(float[] returnValues, int days)
        {
            //Return the array of n days standard deviations for a time period 
            int i,j;
            float[] StDevArr = new float[returnValues.Length - days];
            float[] StDevSubset = new float[days];

            for (i = 0; i<StDevArr.Length; i++)
            {
                
                for (j=0; j<days; j++)
                    StDevSubset[j] = returnValues[j + i];

                StDevArr[i] = StandardDeviation(StDevSubset);
            }
            return StDevArr;
        }
        private static float StandardDeviation(float[] valuesArr)
        {
            float average = valuesArr.Average();
            float sumOfDeviation = 0;
            foreach (float value in valuesArr)
            {
                sumOfDeviation += (value - average) * (value - average);
            }
            float sumOfDeviationAverage = sumOfDeviation / (valuesArr.Length-1);
            return (float)Math.Sqrt(sumOfDeviationAverage);
        }
    }
    
}
