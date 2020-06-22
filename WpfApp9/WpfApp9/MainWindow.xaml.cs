using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Threading;
using System.Windows;

namespace WpfApp9
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public double t0 = 1;
        public double tf = 11;
        public double gamin, gamax, tmin;
        public double tmax = 0;
        public double pmin;                                 // value of Ф(gamma,tau)
        public double gmin1, gmin2;                           // value of gamma 
        public double Tmin1, Tmin2;                            //  value of tau
        public List<double> data = new List<double>();      // value of function fi(t)

        string fileExcel;

        public MainWindow()
        {
            InitializeComponent();
        }
        // function f = g - fi(t) - fi(t+ tau)
        public double function(double t, double T, double g)
        {
            double fi = data[Convert.ToInt32(t * 100)];
            double fi1 = data[Convert.ToInt32((t + T) * 100)];
            return g - fi - fi1;
        }
        // function gamma
        public void gamma()
        {
            double min = 10000000;
            double max = -1000000;
            for (double T = 0; T <= tf - t0; T += 0.01)
            {
                for (double t = t0; t <= tf; t += 0.01)
                {
                    // gamma = fi + fi1
                    if (data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + T) * 100)] > max)
                        max = data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + T) * 100) ];
                    if (data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + T) * 100)] < min)
                        min = data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + T) * 100)];
                }
            }
            gamin = min;            // minumum value of gamma when t in (t0,tf); tau in (0,tf-t0)
            gamax = max;            // maximum value of gamma when t in (t0,tf); tau in (0,tf-t0)
        }

        // value of integral
        public double inte()
        {
            double r, x, T, t, tmin, g;
            double pmin = 70000;
           
                for (T = 0; T <= tf - t0; T += 0.01)
                {
                    double Integral = 0;
                    for (t = t0; t <= tf; t += 0.01)
                    {                        
                        g = data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + T) * 100)];
                        Integral = Integral + Math.Pow(function(t, T, g), 2) * 0.01;
                    }
                        if (Integral < pmin)
                        {
                            pmin = Integral;
                            Tmin1 = T;
                            tmin = t;
                        }
                    
                }
           
            return pmin;
        }
        public double gmin11()
        {
            return data[Convert.ToInt32(tmin * 100)] + data[Convert.ToInt32((tmin + Tmin1) * 100)]; ;
        }
        public double taumin1()
        {
            return Tmin1;
        }
        public double timemax()
        {
            double Max;
            double max1 = -70000;
            for (double t = t0; t < tf; t += 0.01)
            {
                Max = data[Convert.ToInt32(t * 100)];
                if (Max > max1)
                {
                    max1 = Max;
                    tmax = t;
                }
            }
            return tmax;
        }
        // value of min(gamma, tau){max(t)}
        public double minmax()
        {
            double g, T, t;
            double min2 ;
            double Min2 = 70000;
                     
            for (g = gamin; g < gamax; g += 1000)
            {
                for (T = 0; T <= tf - t0; T += 0.01)
                {
                    min2 = Math.Abs(function(timemax(), T, g));
                    if (min2 < Min2)
                    {
                        Min2 = min2;
                        gmin2 = g;
                        Tmin2 = T;
                    }
                }
             }                             
            return Min2;
        }
        
        public double taumin2()
        {
            minmax();
            return Tmin2;
        }
        public double gmin22()
        {
            minmax();
            return gmin2;
        }
        private void btn1_Click(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range Rng;
            fileExcel = @"C:\Users\hp\source\repos\WpfApp9\WpfApp9\obj\Debug\SK_Moschnost.xlsx";
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            // open workbook
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets[1];
            Microsoft.Office.Interop.Excel.Chart xlChart;
            Microsoft.Office.Interop.Excel.Series xlSeries;
            xlApp.Visible = true;
            xlApp.UserControl = true;
            Microsoft.Office.Interop.Excel.Range usedColumn = xlSheet.UsedRange.Columns[2];
            System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

            for (int i = 1; i < strArray.Length; i++)
                data.Add(Convert.ToDouble(strArray[i]));
            
            for (double g = gamin; g < gamax; g += 1000)
            {
                for (double T = 0; T <= tf - t0; T += 0.01)
                {
                    for (double t = t0; t <= tf; t += 0.01)
                    {
                        function(t, T, g);
                    }
                }
            }
            gamma();
            inte();
            taumin1();
            gmin11();
            String sMsg;
            sMsg = "Minimum of integral Ф(gamma,tau): ";
            sMsg = String.Concat(sMsg, inte());
            sMsg = String.Concat(sMsg, " Вт, when minimum of gamma: ");
            sMsg = String.Concat(sMsg, gmin11());
            sMsg = String.Concat(sMsg, ", minimum of tau: ");
            sMsg = String.Concat(sMsg, taumin1());

            MessageBoxResult mes = MessageBox.Show(sMsg, "Caculate and draw graphic?", MessageBoxButton.YesNo);
            if (mes == MessageBoxResult.No)
            {
                Close();
            }
            else
            {
                // Add table headers going cell by cell.
                xlSheet.Cells[1, 4] = "Время [t0,tf], сек.";
                xlSheet.Cells[1, 5] = "gamma";
                

                //AutoFit columns A:B.
                Rng = xlSheet.get_Range("A1:G1");
                Rng.EntireColumn.AutoFit();

                // interval [t0, tf]
                Rng = xlApp.get_Range("D2", "D1002");
                Rng.Formula = "=A101";


                StreamWriter txt = new StreamWriter("testinte1.txt");
                for (double t = t0; t < tf; t += 0.01)
                {
                    txt.WriteLine(data[Convert.ToInt32(t * 100)] + data[Convert.ToInt32((t + taumin1()) * 100)]);
                }
                txt.Close();
                object misvalue = System.Reflection.Missing.Value;
                string[] txtname = System.IO.File.ReadAllLines(@"C:\Users\hp\source\repos\WpfApp9\WpfApp9\bin\Debug\testinte1.txt");
                try
                {
                    for (int i = 0; i <= txtname.Length; i++)
                    {
                        xlSheet.Cells[5][i + 2] = txtname[i];
                    }
                    Thread.Sleep(3000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception" + ex);
                }
                // add a chart for the selected data
                xlWorkBook = (Microsoft.Office.Interop.Excel.Workbook)xlSheet.Parent;
                xlChart = (Microsoft.Office.Interop.Excel.Chart)xlWorkBook.Charts.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                // use the ChartWizard to create a new chart from the select data
                xlSeries = (Microsoft.Office.Interop.Excel.Series)xlChart.SeriesCollection(1);
                xlSeries.XValues = xlSheet.get_Range("E2:E1002");
               }
        }
        private void btn2_Click(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Range Rng;
            fileExcel = @"C:\Users\hp\source\repos\WpfApp9\WpfApp9\obj\Debug\SK_Moschnost.xlsx";
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            // open workbook
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileExcel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets[1];
            Microsoft.Office.Interop.Excel.Chart xlChart;
            Microsoft.Office.Interop.Excel.Series xlSeries;
            xlApp.Visible = true;
            xlApp.UserControl = true;
            Microsoft.Office.Interop.Excel.Range usedColumn = xlSheet.UsedRange.Columns[2];
            System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
            string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

            for (int i = 1; i < strArray.Length; i++)
                data.Add(Convert.ToDouble(strArray[i]));

            for (double g = gamin; g < gamax; g += 1000)
            {
                for (double T = 0; T <= tf - t0; T += 0.01)
                {
                    for (double t = t0; t <= tf; t += 0.01)
                    {
                        function(t, T, g);                        
                    }
                }
            }
            gamma();
            minmax();
            timemax();
            taumin2();
            gmin22();
            String sMsg;
            sMsg = "Minimum of maximum |f(gamma,tau)|: ";
            sMsg = String.Concat(sMsg, minmax());
            sMsg = String.Concat(sMsg, " Вт, when minimum of gamma: ");
            sMsg = String.Concat(sMsg, gmin22());
            sMsg = String.Concat(sMsg, ", minimum of tau: ");
            sMsg = String.Concat(sMsg, taumin2());
            sMsg = String.Concat(sMsg, ", maximum of time: ");
            sMsg = String.Concat(sMsg, timemax());

            MessageBoxResult mes = MessageBox.Show(sMsg, "Caculate and draw graphic?", MessageBoxButton.YesNo);
            if (mes == MessageBoxResult.No)
            {
                Close();
            }
            else
            {
                //Add table headers going cell by cell.
                xlSheet.Cells[1, 4] = "Время [t0,tf], сек.";
                xlSheet.Cells[1, 5] = "|f(gamma,tau)|";
               

                //AutoFit columns A:B.
                Rng = xlSheet.get_Range("A1:G1");
                Rng.EntireColumn.AutoFit();

                // interval [t0, tf]
                Rng = xlApp.get_Range("D2", "D1002");
                Rng.Formula = "=A101";


                StreamWriter txt = new StreamWriter("testinte2.txt");
                for (double t = t0; t < tf; t += 0.01)
                {
                    txt.WriteLine(Math.Abs(gmin22() - data[Convert.ToInt32(t * 100)] - data[Convert.ToInt32((t + taumin2()) * 100)]));
                }
                txt.Close();
                object misvalue = System.Reflection.Missing.Value;
                string[] txtname = System.IO.File.ReadAllLines(@"C:\Users\hp\source\repos\WpfApp9\WpfApp9\bin\Debug\testinte2.txt");
                try
                {
                    for (int i = 0; i <= txtname.Length; i++)
                    {
                        xlSheet.Cells[5][i + 2] = txtname[i];
                    }
                    Thread.Sleep(3000);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception" + ex);
                }
                // add a chart for the selected data
                xlWorkBook = (Microsoft.Office.Interop.Excel.Workbook)xlSheet.Parent;
                xlChart = (Microsoft.Office.Interop.Excel.Chart)xlWorkBook.Charts.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                // use the ChartWizard to create a new chart from the select data
                xlSeries = (Microsoft.Office.Interop.Excel.Series)xlChart.SeriesCollection(1);
                xlSeries.XValues = xlSheet.get_Range("E2:E1002");
                }
            }

        }
    }

