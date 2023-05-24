using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ClassTesterFinal
{
    public class CalcCodeClass
    {
        private double seriennummer_kd;
        private double codetoolic;
        private double codereign;
        private double codeschltprgm;
        private double codetlgrmsnd;
        private double codetelicm;



        public double Seriennummer_kd
        {
            get { return seriennummer_kd; }
            set { seriennummer_kd = value; }
        }


        public double CodeToolIc
        {
            get { return codetoolic; }
            set { codetoolic = value; }
        }


        public double CodeEreign
        {
            get { return codereign; }
            set { codereign = value; }
        }


        public double CodeSchltPrgm
        {
            get { return codeschltprgm; }
            set { codeschltprgm = value; }
        }


        public double CodeTlgrmsnd
        {
            get { return codetlgrmsnd; }
            set { codetlgrmsnd = value; }
        }


        public double CodeTelicm
        {
            get { return codetelicm; }
            set { codetelicm = value; }
        }


        //______________________________________FreischaltCodeBerechnen___________________________________________________
        public void FreischaltCodeBerechnen(string seriennummer_kd_str)
        {
            int seriennummer_kd;
            if (int.TryParse(seriennummer_kd_str, out seriennummer_kd))
            {
                codetoolic = Math.Round(seriennummer_kd / 3.0, MidpointRounding.ToEven) + 123456 +
                                Math.Round(seriennummer_kd / 5.0, MidpointRounding.ToEven) +
                                Math.Round(seriennummer_kd * 3 / 13.0, MidpointRounding.ToEven) +
                                Math.Round(seriennummer_kd * 3 / 7.0, MidpointRounding.ToEven);


                codereign = Math.Round(seriennummer_kd / 5.0, MidpointRounding.ToEven) * 3 + 76855 +
                                Math.Round(seriennummer_kd / 7.0, MidpointRounding.ToEven) * 6 +
                                Math.Round(seriennummer_kd / 13.0, MidpointRounding.ToEven) * 5 +
                                Math.Round((double)seriennummer_kd / 11, MidpointRounding.ToEven) * 7;

                codeschltprgm = Math.Round(seriennummer_kd / 15.0, MidpointRounding.ToEven) * 11 + 83125 +
                                Math.Round(seriennummer_kd / 5.0, MidpointRounding.ToEven) * 9 +
                                Math.Round(seriennummer_kd / 4.0, MidpointRounding.ToEven) * 7 +
                                Math.Round(seriennummer_kd / 16.0, MidpointRounding.ToEven) * 9;

                codetlgrmsnd = Math.Round(seriennummer_kd / 2.0, MidpointRounding.ToEven) + 369825 +
                                Math.Round(seriennummer_kd / 7.0, MidpointRounding.ToEven) * 3 +
                                Math.Round(seriennummer_kd / 25.0, MidpointRounding.ToEven) * 13 +
                                Math.Round((double)seriennummer_kd / 13, MidpointRounding.ToEven) * 10;

                codetelicm = Math.Round(seriennummer_kd / 3.0, MidpointRounding.ToEven) + 234567 +
                                Math.Round(seriennummer_kd / 5.0, MidpointRounding.ToEven) +
                                Math.Round(seriennummer_kd * 3 / 12.0, MidpointRounding.ToEven) +
                                Math.Round(seriennummer_kd * 3 / 8.0, MidpointRounding.ToEven) +
                                Math.Round((double)seriennummer_kd * 2 / 3, MidpointRounding.ToEven);

            }
            else
            {
                MessageBox.Show("Bitte geben Sie eine gültige Seriennummer ein.");
            }
            //__________________________________________FreischaltCodeBerechnen___________________________________________________          
        }
    }
}
