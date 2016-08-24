using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using HistoBin = System.Tuple<string, ExceLint.Scope.SelectID, double>;
using FreqTable = System.Collections.Generic.Dictionary<System.Tuple<string, ExceLint.Scope.SelectID, double>, int>;
using Color = System.Drawing.Color;
using Distribution = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<ExceLint.Scope.SelectID, System.Collections.Generic.Dictionary<double, Microsoft.FSharp.Collections.FSharpSet<AST.Address>>>>;
using XYInfo = System.Collections.Generic.Dictionary<System.Tuple<double,double>,AST.Address>;

namespace ExceLintUI
{
    public partial class SpectralPlot : Form
    {
        ExceLint.Scope.SelectID[] cs;
        string[] fs;
        ExceLint.ErrorModel m;
        bool drawn = false;
        ToolTip tooltip = new ToolTip();
        Distribution d;
        XYInfo xyinfo;

        public SpectralPlot(ExceLint.ErrorModel model)
        {
            InitializeComponent();

            // init model
            m = model;

            // init distribution
            d = m.Distribution;

            // init mouseover info
            xyinfo = new XYInfo();

            // init combo box data sources
            cs = m.FrequencyTable.Keys.Select((HistoBin h) => h.Item2).Distinct().ToArray();
            fs = m.Features;

            // init combo boxes;
            var selectorNames = cs.Select((ExceLint.Scope.SelectID s) => ExceLint.Scope.Selector.ToPretty(s)).ToArray();
            comboCondition.DataSource = selectorNames;
            comboFeature.DataSource = fs;
        }


        private void SpectralPlot_Load(object sender, EventArgs e)
        {
            drawPlot();
        }

        private void drawPlot()
        {
            // remove any pre-existing series
            chart1.Series.Clear();

            // get combo selections
            // choose the first one if nothing is selected
            ExceLint.Scope.SelectID s = cs[comboCondition.SelectedIndex == -1 ? 0 : comboCondition.SelectedIndex];
            string f = fs[comboFeature.SelectedIndex == -1 ? 0 : comboFeature.SelectedIndex];

            // draw plot
            drawBins(f, s);

            // draw
            chart1.Invalidate();

            drawn = true;
        }

        private void drawBins(string feature, ExceLint.Scope.SelectID condition)
        {
            // clear xyinfo
            xyinfo.Clear();

            // which subset of bins to plot?
            var bins = m.FrequencyTable.Keys.Where((HistoBin h) => h.Item1 == feature && h.Item2 == condition).ToArray();

            // find dynamic range of features
            var fMin = bins.Select((HistoBin h) => h.Item3).Min();
            var fMax = bins.Select((HistoBin h) => h.Item3).Max();

            // plot them
            int i = 0;
            foreach (HistoBin h in bins)
            {
                drawBin(feature, condition, h, m.FrequencyTable, getColor(h.Item3, fMin, fMax));
            }
        }

        private static Color getColor(double featureValue, double featureMin, double featureMax)
        {
            // inspired by function found here: https://en.wikipedia.org/wiki/HSL_and_HSV#From_HSV

            double normedFeature;
            if ((featureMax - featureMin) == 0)
            {
                normedFeature = 0;
            } else
            {
                normedFeature = (featureValue - featureMin) / (featureMax - featureMin);
            }
            
            double hue = normedFeature * 360;
            double saturation = 1;
            double lightness = 1;
            double chroma = saturation * lightness;
            double huePrime = hue / 60;
            double X = chroma * (1 - Math.Abs(huePrime % 2 - 1));
            double R1 = 0, G1 = 0, B1 = 0;

            if (huePrime >= 0 && huePrime < 1)
            {
                R1 = chroma;
                G1 = X;
                B1 = 0;
            } else if (huePrime >= 1 && huePrime < 2)
            {
                R1 = X;
                G1 = chroma;
                B1 = 0;
            } else if (huePrime >= 2 && huePrime < 3)
            {
                R1 = 0;
                G1 = chroma;
                B1 = X;
            } else if (huePrime >= 3 && huePrime < 4)
            {
                R1 = 0;
                G1 = X;
                B1 = chroma;
            } else if (huePrime >= 4 && huePrime < 5)
            {
                R1 = X;
                G1 = 0;
                B1 = chroma;
            } else
            {
                R1 = chroma;
                G1 = 0;
                B1 = X;
            }

            double m = lightness - chroma;

            double R = R1 + m,
                   G = G1 + m,
                   B = B1 + m;

            System.Diagnostics.Debug.Assert(R >= 0 && R <= 1);
            System.Diagnostics.Debug.Assert(G >= 0 && G <= 1);
            System.Diagnostics.Debug.Assert(B >= 0 && B <= 1);

            int r = Convert.ToInt32(R * 255),
                g = Convert.ToInt32(G * 255),
                b = Convert.ToInt32(B * 255);

            return Color.FromArgb(255, r, g, b);
        }

        private void drawBin(string feature, ExceLint.Scope.SelectID sid, HistoBin h, FreqTable freqtable, Color c)
        {
            var binName = h.Item3.ToString();
            var hashValue = h.Item3;
            var addresses = d[feature][sid][hashValue].ToArray();
            var count = freqtable[h];

            // create series for scatterplot
            var series = new Series
            {
                Name = binName,
                Color = c,
                IsVisibleInLegend = false,
                IsXValueIndexed = false,
                ChartType = SeriesChartType.Point,
                AxisLabel = binName
            };

            // add series to chart
            chart1.Series.Add(series);

            // generate data
            var xData = new double[count];
            var yData = new double[count];
            for (int i = 1; i <= count; i++)
            {
                xData[i-1] = hashValue;
                yData[i-1] = i;
            }

            //// generate map for mouseover
            //for (int i = 0; i < count; i++)
            //{
            //    var xy = new Tuple<double, double>(xData[i], yData[i]);
            //    xyinfo.Add(xy, addresses[i]);
            //}

            // bind data to plot
            chart1.Series[binName].Points.DataBindXY(xData, yData);
        }

        private void comboCondition_SelectedIndexChanged(object sender, EventArgs e)
        {
            drawPlot();
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            if (!drawn)
            {
                return;
            }

            // get combo selections
            // choose the first one if nothing is selected
            ExceLint.Scope.SelectID s = cs[comboCondition.SelectedIndex == -1 ? 0 : comboCondition.SelectedIndex];
            string f = fs[comboFeature.SelectedIndex == -1 ? 0 : comboFeature.SelectedIndex];

            var pos = e.Location;

            var results = chart1.HitTest(pos.X, pos.Y, false,
                                    ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 2 &&
                            Math.Abs(pos.Y - pointYPixel) < 2)
                        {
                            //var xy = new Tuple<double, double>(pointXPixel, pointYPixel);
                            //var addr = xyinfo[xy];

                            tooltip.Show("hash = " + prop.XValue,
                                         //+ "\r\n" + addr.A1FullyQualified(),
                                         this.chart1,
                                         pos.X,
                                         pos.Y - 15
                                        );
                        }
                    }
                }
            }
        }
    }
}
