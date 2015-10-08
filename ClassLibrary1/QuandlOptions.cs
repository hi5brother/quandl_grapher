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
        public static Double BrennerEstimate2(double callPrice, double underlying,double expiry)
        {
            //http://quant.stackexchange.com/a/7763
            //http://www.eecs.harvard.edu/~parkes/cs286r/spring08/reading3/chambers.pdf
            //Expiry in years
            //For at the money 
            double vol;

            vol = Math.Sqrt(2 * Math.PI / expiry) * (callPrice / underlying);

            return vol;
        }
        public static double IterativeBS(double price, double underlying, double strike, double expiry, double rate, string cp)
        {
            //Newton-Ralpson iterative method

            BlackScholes option = new BlackScholes(price, underlying,  strike, expiry, rate, cp);

            do
            {
                option.CalculateRiskProbabilities();

                option.vol = option.vol - (option.CalculateOptionValue() - price) / option.CalculateVega();

            } while ((option.CalculateOptionValue() - price) > 0.01);
            return option.vol;
        }
        struct BlackScholes
        {
            //SEE: http://www.fincad.com/resources/resource-library/wiki/black-scholes-model
            //SEE: http://finance.bi.no/~bernt/gcc_prog/algoritms_v1/algoritms/node8.html
            
            public double optionPrice, underlying, strike, expiry, rate;
            public string call_put;
            private double d1, d2;
            double sqrtTime;
            public double initVol, vol;
            
            public BlackScholes(double p_optionPrice, 
                                double p_underlying, 
                                double p_strike, 
                                double p_expiry, 
                                double p_rate,
                                string p_call_put)
            {
                optionPrice = p_optionPrice;
                underlying = p_underlying;
                strike = p_strike;
                expiry = p_expiry;      //In Years
                rate = p_rate;
                call_put = p_call_put;
                
                initVol = BrennerEstimate(optionPrice, expiry, strike);
                
                sqrtTime = Math.Sqrt(p_expiry);

                d1 = Math.Log(underlying/strike) + expiry * ((rate) + 0.5 * initVol * initVol);
                d1 = d1 / (initVol * Math.Sqrt(expiry));

                d2 = d1 - initVol * Math.Sqrt(expiry);

                vol = initVol;
            }
            public void CalculateRiskProbabilities()
            {
                d1 = Math.Log(underlying / strike) + expiry * ((rate) + 0.5 * vol *vol);
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
                double value = 0;
                if (call_put == "c")
                    value = underlying * NDistCumul(d1) - strike * Math.Exp(-rate * expiry) * NDistCumul(d2);
                else if (call_put == "p")
                    value = strike * Math.Exp(-rate * expiry) * NDistCumul(-d2) - underlying * NDistCumul(-d1);

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
