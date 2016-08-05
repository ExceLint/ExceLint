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
                Color = System.Drawing.Color.Green,
                IsVisibleInLegend = false,
                IsXValueIndexed = true,
                ChartType = SeriesChartType.Point
            };

            // plot
            var rng = new Random();
            for (int i = 0; i < 5000; i++)
            {
                series1.Points.AddXY(rng.NextDouble(), rng.NextDouble());
            }

            // draw
            chart1.Invalidate();
        }
    }
}
