using System;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace ExceLintUI
{
    public partial class SpectralPlot : Form
    {
        public SpectralPlot(ExceLint.ErrorModel model)
        {
            InitializeComponent();
        }

        private void SpectralPlot_Load(object sender, EventArgs e)
        {
            // remove any pre-existing series
            chart1.Series.Clear();

            // create series for scatterplot
            var series1 = new Series
            {
                Name = "points",
                Color = System.Drawing.Color.Black,
                IsVisibleInLegend = false,
                IsXValueIndexed = true,
                ChartType = SeriesChartType.Point
            };

            // add series to chart
            chart1.Series.Add(series1);

            // generate data
            var xData = new double[5000];
            var yData = new double[5000];
            var rng = new Random();
            for (int i = 0; i < 5000; i++)
            {
                xData[i] = rng.NextDouble();
                yData[i] = rng.NextDouble();
            }

            // bind data to plot
            chart1.Series["points"].Points.DataBindXY(xData, yData);

            // draw
            chart1.Invalidate();
        }
    }
}
