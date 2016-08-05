﻿using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using HistoBin = System.Tuple<string, ExceLint.Scope.SelectID, double>;
using FreqTable = System.Collections.Generic.Dictionary<System.Tuple<string, ExceLint.Scope.SelectID, double>, int>;
using Color = System.Drawing.Color;

namespace ExceLintUI
{
    public partial class SpectralPlot : Form
    {
        ExceLint.Scope.Selector[] cs;
        string[] fs;
        ExceLint.ErrorModel m;

        public SpectralPlot(ExceLint.ErrorModel model)
        {
            InitializeComponent();

            // init model
            m = model;

            // init combo box data sources
            cs = model.Scopes;
            fs = model.Features;

            // init combo boxes;
            var selectorNames = cs.Select((ExceLint.Scope.Selector s) => s.ToString()).ToArray();
            comboCondition.DataSource = selectorNames;
            comboFeature.DataSource = fs;
        }

        private void SpectralPlot_Load(object sender, EventArgs e)
        {
            // remove any pre-existing series
            chart1.Series.Clear();

            // get combo selections
            ExceLint.Scope.Selector s = cs[comboCondition.SelectedIndex];
            string f = fs[comboFeature.SelectedIndex];

            // draw plot
            drawPlot(f, s);

            // draw
            chart1.Invalidate();
        }

        private void drawPlot(string feature, ExceLint.Scope.Selector condition)
        {
            // which subset of bins to plot?
            var bins = m.FrequencyTable.Keys.Where((HistoBin h) => h.Item1 == feature && h.Item2.IsKind == condition).ToArray();

            // find dynamic range of features
            var fMin = bins.Select((HistoBin h) => h.Item3).Min();
            var fMax = bins.Select((HistoBin h) => h.Item3).Max();

            // plot them
            int i = 0;
            foreach (HistoBin h in bins)
            {
                drawBin(h, m.FrequencyTable, getColor(h.Item3, fMin, fMax));
            }
        }

        private static Color getColor(double featureValue, double featureMin, double featureMax)
        {
            // inspired by function found here: https://en.wikipedia.org/wiki/HSL_and_HSV#From_HSV
            double normedFeature = (featureValue - featureMin) / (featureMax - featureMin);
            double hue = normedFeature * 360;
            double saturation = 1;
            double lightness = 1;
            double chroma = saturation * lightness;
            double huePrime = hue / 60;
            double X = chroma * (1 - (huePrime % 2 - 1));
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

            int R = Convert.ToInt32(R1 + m),
                G = Convert.ToInt32(G1 + m),
                B = Convert.ToInt32(B1 + m);

            return Color.FromArgb(R, G, B);
        }

        private void drawBin(HistoBin h, FreqTable freqtable, Color c)
        {
            var binName = h.Item3.ToString();
            var hashValue = h.Item3;
            var count = freqtable[h];

            // create series for scatterplot
            var series1 = new Series
            {
                Name = binName,
                Color = c,
                IsVisibleInLegend = false,
                IsXValueIndexed = false,
                ChartType = SeriesChartType.Point
            };

            // add series to chart
            chart1.Series.Add(series1);

            // generate data
            var xData = new double[count];
            var yData = new double[count];
            for (int i = 1; i <= count; i++)
            {
                xData[i-1] = hashValue;
                yData[i-1] = i;
            }

            // bind data to plot
            chart1.Series[binName].Points.DataBindXY(xData, yData);
        }
    }
}
