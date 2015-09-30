using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace QuandlProcess
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("QuandlOptions")]
    //Interface with VBA as
    //Dim lib As Object: Set lib = CreateObject("QuandlProcess")
    //See: http://mikejuniperhill.blogspot.ca/2014/03/interfacing-c-and-vba-with-exceldna-no.html

    public class QuandlData
    {
        private ParsedData quandlObject;

        public Boolean InstantiateQuandl(string quandlDatabase, string datacodeParams, int datapointsNumber, string frequency)
        {
            this.quandlObject = QuandlProcess.GrabData(quandlDatabase, datacodeParams, datapointsNumber, frequency);
            return true;
        }
        public Double HistoricalVol(int days)
        {
            //Find historical vol over last ___ days
            double avgChange = 0;
            int numOfDays;
            int numOfIntervals;

            double histVol = 0;
            double differenceVal;

            float[] closeData = this.quandlObject.ReturnFloatValues("Close");
            numOfDays = closeData.Length;

            if (days > numOfDays)
                numOfIntervals = days;
            else
                numOfIntervals = numOfDays - 1;

            float[] rateOfReturn = new float[numOfIntervals];

            //First loop to calculate average change
            for (int i = 0; i < numOfIntervals; i++)
            {
                rateOfReturn[i] = (float)Math.Log((double)closeData[i] / closeData[i + 1]);
                avgChange += rateOfReturn[i];
            }
            avgChange = avgChange / numOfIntervals;

            //Second loop to calculate difference between 
            for (int i = 0; i < numOfIntervals; i++)
            {
                differenceVal = (rateOfReturn[i] - avgChange);
                histVol += differenceVal * differenceVal;
            }
            histVol = histVol / (numOfIntervals - 1);

            histVol = Math.Sqrt(histVol);

            //Annualize the volatility
            histVol = histVol * Math.Sqrt(252);

            histVol = histVol * 100;

            return histVol;
        }
        public static Double BrennerEstimate(double callPrice, double expiry, double strike)
        {
            //http://quant.stackexchange.com/a/7763
            //http://www.eecs.harvard.edu/~parkes/cs286r/spring08/reading3/chambers.pdf
            //Expiry in years
            //For at the money 
            double vol;

            vol = Math.Sqrt(2 * Math.PI / expiry) * (callPrice / strike);

            return vol;

        }
        public Double BrennerEstimate2(double callPrice, double expiry, double strike)
        {
            //http://quant.stackexchange.com/a/7763
            //http://www.eecs.harvard.edu/~parkes/cs286r/spring08/reading3/chambers.pdf
            //Expiry in years
            //For at the money 
            double vol;

            vol = Math.Sqrt(2 * Math.PI / expiry) * (callPrice / strike);

            return vol;
        }
        public double IterativePricer(double price, double underlying, double strike, double expiry, double rate)
        {

            BlackScholes test = new BlackScholes(price, underlying,  strike, expiry, rate);
            
            double vol;
            for (int i = 0; i < 100; i++ )
            {
                
                test.CalculateRiskProbabilities();

                vol = test.vol - (test.CalculateOptionValue() - price) / test.CalculateVega() ;

            }
            return test.vol;
        }
        struct BlackScholes
        {
            //SEE: http://www.fincad.com/resources/resource-library/wiki/black-scholes-model
            //SEE: http://finance.bi.no/~bernt/gcc_prog/algoritms_v1/algoritms/node8.html
            
            public double optionPrice, underlying, strike, expiry, rate;
            private double d1, d2;
            double sqrtTime;
            public double initVol, vol;
            
            public BlackScholes(double p_optionPrice, 
                                double p_underlying, 
                                double p_strike, 
                                double p_expiry, 
                                double p_rate)
            {
                optionPrice = p_optionPrice;
                underlying = p_underlying;
                strike = p_strike;
                expiry = p_expiry;      //In Years
                rate = p_rate;
                
                initVol = BrennerEstimate(optionPrice, expiry, strike);
                
                sqrtTime = Math.Sqrt(p_expiry);

                d1 = Math.Log(underlying/strike) + expiry * ((rate) + 0.5 * initVol * initVol);
                d1 = d1 / (initVol * Math.Sqrt(expiry));

                d2 = d1 - initVol * Math.Sqrt(expiry);

                vol = initVol;
            }
            public void CalculateRiskProbabilities()
            {
                d1 = Math.Log(underlying / strike) + expiry * ((rate) + 0.5 * vol * initVol);
                d1 = d1 / (vol * sqrtTime);

                d2 = d1 - vol * sqrtTime;
            }
            public double CalculateVega()
            {
                //SEE: http://finance.bi.no/~bernt/gcc_prog/algoritms_v1/algoritms/node8.html
                double vega;
                vega = underlying * sqrtTime * NDistCumul(d1);
                return vega;
            }
            public double CalculateOptionValue()
            {
                double value;
                value = underlying * NDistCumul(d1) * Math.Exp(-expiry) - strike * Math.Exp(-rate * expiry) * NDistCumul(d2);
                return value;
            }
            private double NDistCumul(double val)
            {
                //SEE https://software.intel.com/en-us/node/531898
                return 0.5 + 0.5 * Erf(val / Math.Sqrt(2));
                
            }
            private double Erf(double val)
            {
                //SEE http://www.johndcook.com/blog/csharp_erf/
                // constants
                double a1 = 0.254829592;
                double a2 = -0.284496736;
                double a3 = 1.421413741;
                double a4 = -1.453152027;
                double a5 = 1.061405429;
                double p = 0.3275911;

                // Save the sign of x
                int sign = 1;
                if (val < 0)
                    sign = -1;
                val = Math.Abs(val);

                // A&S formula 7.1.26
                double t = 1.0 / (1.0 + p * val);
                double y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.Exp(-val * val);

                return sign * y;
            }
        }
    }
}
