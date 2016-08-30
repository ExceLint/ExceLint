using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using HistoBin = System.Tuple<string, ExceLint.Scope.SelectID, double>;
using Color = System.Drawing.Color;
using Distribution = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<ExceLint.Scope.SelectID, System.Collections.Generic.Dictionary<double, Microsoft.FSharp.Collections.FSharpSet<AST.Address>>>>;
using XYInfo = System.Collections.Generic.Dictionary<double,System.Collections.Generic.Dictionary<double, System.Tuple<AST.Address,bool,string>>>;
using Microsoft.FSharp.Collections;
using System.Collections.Generic;

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
            var bins = d[feature][condition];

            // find dynamic range of features
            var fMin = bins.Select((pair) => pair.Key).Min();
            var fMax = bins.Select((pair) => pair.Key).Max();

            // store all anomalies in separate set
            var anoms = new Dictionary<AST.Address, double>();
            foreach (var pair in m.rankByFeatureSum().Take(m.Cutoff + 1))
            {
                anoms.Add(pair.Key, pair.Value);
            }

            // plot them, sans anomalies
            foreach (var pair in bins)
            {
                var hashValue = pair.Key;
                var addresses = pair.Value;
                var sxy = drawBin(addresses, hashValue, Color.Black, xyinfo, anoms, feature);
                var s = sxy.Item1;
                xyinfo = sxy.Item2;
                chart1.Series.Add(s);
            }

            // finally, plot anomalies, one bin at a time
            var mtanoms = new Dictionary<AST.Address, double>();
            var anom_grps = anoms.GroupBy(pair => pair.Value).ToDictionary(grouping => grouping.Key, grouping => grouping.Select(pair => pair.Key));
            foreach (var pair in anom_grps)
            {
                var hashValue = pair.Key;
                var addrs = pair.Value;
                var addresses = new FSharpSet<AST.Address>(addrs);
                var sxy = drawBin(addresses, hashValue, getColor(hashValue, fMin, fMax), xyinfo, mtanoms, feature);
                var s = sxy.Item1;
                xyinfo = sxy.Item2;
                chart1.Series.Add(s);
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

        private static Tuple<Series,XYInfo> drawBin(FSharpSet<AST.Address> addresses, double hashValue, Color c, XYInfo xyinfo, Dictionary<AST.Address, double> anoms, string feature)
        {
            var binName = hashValue.ToString();

            bool isAnom = anoms.Count == 0;

            // create series for scatterplot
            var series = new Series
            {
                Name = isAnom ? binName : binName + "anom",
                Color = c,
                IsVisibleInLegend = false,
                IsXValueIndexed = false,
                ChartType = SeriesChartType.Point,
                AxisLabel = binName
            };

            // generate data
            var xData = new double[addresses.Count];
            var yData = new double[addresses.Count];

            var addrs = addresses.ToArray();
            for (int i = 0; i < addrs.Length; i++)
            {
                var addr = addrs[i];
                double y = 0;
                if (xyinfo.ContainsKey(hashValue))
                {
                    // find the height of the bin and add one
                    var bin = xyinfo[hashValue].Select(pair => pair.Key);
                    var max_y = bin.Count() > 0 ? bin.Max() : 0;
                    y = max_y + 1;
                } else
                {
                    xyinfo.Add(hashValue, new Dictionary<double, Tuple<AST.Address,bool,string>>());
                }

                if (!anoms.ContainsKey(addr))
                {
                    xData[i] = hashValue;
                    yData[i] = y;
                    xyinfo[hashValue].Add(y, new Tuple<AST.Address,bool,string>(addr, isAnom, feature));
                }
            }

            // bind data to plot
            series.Points.DataBindXY(xData, yData);

            return new Tuple<Series,XYInfo>(series, xyinfo);
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
                        if (Math.Abs(pos.X - pointXPixel) <= 1 &&
                            Math.Abs(pos.Y - pointYPixel) <= 1)
                        {
                            var addr = xyinfo[prop.XValue][prop.YValues[0]].Item1;
                            var isAnom = xyinfo[prop.XValue][prop.YValues[0]].Item2;
                            var feature = xyinfo[prop.XValue][prop.YValues[0]].Item3;

                            var ttstr = "hash: " + prop.XValue
                                         + "\r\n" + addr.A1Worksheet() + "!" + addr.A1Local();
                            if (m.Fixes != null && isAnom)
                            {
                                var fixes = m.Fixes.Value;
                                var fix = fixes[addr][feature];
                                ttstr = ttstr + "\r\nfix: " + fix;
                                
                            }

                            tooltip.Show(ttstr,
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
