using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.DataVisualization.Charting;
using XsPDF.Drawing;
using XsPDF.Pdf;

namespace XsPDF.Demo
{
    class MSChartDemo
    {
        public static void InsertMSChartToPDF()
        {
            // Create a new PDF document.
            PdfDocument document = new PdfDocument();

            // Create an empty page in this document.
            PdfPage page = document.AddPage();

            // Obtain an XGraphics object to render to
            XGraphics g = XGraphics.FromPdfPage(page);

            // Create range-column chart stream object
            Stream chartStream = CreateMSRangeColumnChart();

            // Convert chart stream to XImage object
            XImage chartImage = XImage.FromStream(chartStream);

            // Insert the chart image to pdf page in any position
            g.DrawImage(chartImage, 50, 50);

            // Save and show the document
            document.Save("MSRangeColumnChart.pdf");
            Process.Start("MSRangeColumnChart.pdf");
        }

        public static Stream CreateMSRangeColumnChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Bind series data to chart
            // If use range chart type, need bind the series to chart first
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            DateTime currentData = DateTime.Now.Date;

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.RangeColumn;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.CustomProperties = "PointWidth=0.7";
            series1.Points.AddXY(2, currentData, currentData.AddDays(4));
            series1.Points.AddXY(4, currentData.AddDays(5), currentData.AddDays(7));
            series1.Points.AddXY(3, currentData.AddDays(8), currentData.AddDays(10));
            series1.Points.AddXY(1, currentData.AddDays(12), currentData.AddDays(15));
            series1.Points.AddXY(4, currentData.AddDays(17), currentData.AddDays(22));

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.RangeColumn;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.CustomProperties = "DrawSideBySide=false, PointWidth=0.2";
            series2.Points.AddXY(1, currentData, currentData.AddDays(4));
            series2.Points.AddXY(2, currentData.AddDays(5), currentData.AddDays(7));
            series2.Points.AddXY(3, currentData.AddDays(8), currentData.AddDays(10));
            series2.Points.AddXY(5, currentData.AddDays(12), currentData.AddDays(13));
            series2.Points.AddXY(2, currentData.AddDays(12), currentData.AddDays(16));

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSRangeBarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Bind series data to chart
            // If use range chart type, need bind the series to chart first
            chart.Series.Add(series1);
            chart.Series.Add(series2);
                        
            DateTime currentData = DateTime.Now.Date;

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.RangeBar;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.CustomProperties = "PointWidth=0.7";
            series1.Points.AddXY(2, currentData, currentData.AddDays(4));
            series1.Points.AddXY(4, currentData.AddDays(5), currentData.AddDays(7));
            series1.Points.AddXY(3, currentData.AddDays(8), currentData.AddDays(10));
            series1.Points.AddXY(1, currentData.AddDays(12), currentData.AddDays(15));
            series1.Points.AddXY(4, currentData.AddDays(17), currentData.AddDays(22));

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.RangeBar;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.CustomProperties = "DrawSideBySide=false, PointWidth=0.2";
            series2.Points.AddXY(1, currentData, currentData.AddDays(4));
            series2.Points.AddXY(2, currentData.AddDays(5), currentData.AddDays(7));
            series2.Points.AddXY(3, currentData.AddDays(8), currentData.AddDays(10));
            series2.Points.AddXY(5, currentData.AddDays(12), currentData.AddDays(13));
            series2.Points.AddXY(2, currentData.AddDays(12), currentData.AddDays(16));

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSSplineRangeChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Bind series data to chart
            // If use range chart type, need bind the series to chart first
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Input Data Point
            double[] yValue11 = { 56, 74, 45, 59, 34, 87, 50, 87, 64, 34 };
            double[] yValue12 = { 42, 65, 30, 42, 25, 47, 40, 70, 34, 20 };
            double[] yValue21 = { 26, 54, 35, 79, 64, 37, 70, 67, 34, 74 };
            double[] yValue22 = { 12, 6, 23, 34, 15, 27, 60, 30, 24, 50 };

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.SplineRange;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.DataBindY(yValue11, yValue12);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.SplineRange;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.DataBindY(yValue21, yValue22);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

          
            return ms;
        }

        public static Stream CreateMSRangeChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Bind series data to chart
            // If use range chart type, need bind the series to chart first
            chart.Series.Add(series1);
            chart.Series.Add(series2); 
                        
            // Input Data Point
            double[] yValue11 = { 56, 74, 45, 59, 34, 87, 50, 87, 64, 34 };
            double[] yValue12 = { 42, 65, 30, 42, 25, 47, 40, 70, 34, 20 };
            double[] yValue21 = { 26, 54, 35, 79, 64, 37, 70, 67, 34, 74 };
            double[] yValue22 = { 12, 6, 23, 34, 15, 27, 60, 30, 24, 50 };

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Range;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.DataBindY(yValue11, yValue12);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Range;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.DataBindY(yValue21, yValue22);
                        
            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSPyramidChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 39);
            DataPoint dataPoint2 = new DataPoint(0, 18);
            DataPoint dataPoint3 = new DataPoint(0, 15);
            DataPoint dataPoint4 = new DataPoint(0, 12);
            DataPoint dataPoint5 = new DataPoint(0, 5);


            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(64, 64, 64, 64);
            series1.ChartType = SeriesChartType.Pyramid;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(180, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);


            // Bind series data to chart
            chart.Series.Add(series1);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            //chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSFunnelChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 39);
            DataPoint dataPoint2 = new DataPoint(0, 18);
            DataPoint dataPoint3 = new DataPoint(0, 15);
            DataPoint dataPoint4 = new DataPoint(0, 12);
            DataPoint dataPoint5 = new DataPoint(0, 5);


            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(64, 64, 64, 64);
            series1.ChartType = SeriesChartType.Funnel;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(180, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);


            // Bind series data to chart
            chart.Series.Add(series1);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            //chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSBubbleChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 6);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Bubble;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.MarkerSize = 10;
            series1.MarkerStyle = MarkerStyle.Circle;           
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Bubble;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.MarkerSize = 10;
            series2.MarkerStyle = MarkerStyle.Square;
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSPointChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 6);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Point;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.MarkerSize = 10;
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Point;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.MarkerSize = 10;
            series2.MarkerStyle = MarkerStyle.Square;
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSDoughnutChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 39);
            DataPoint dataPoint2 = new DataPoint(0, 18);
            DataPoint dataPoint3 = new DataPoint(0, 15);
            DataPoint dataPoint4 = new DataPoint(0, 12);
            DataPoint dataPoint5 = new DataPoint(0, 5);


            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(64, 64, 64, 64);
            series1.ChartType = SeriesChartType.Doughnut;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(180, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);


            // Bind series data to chart
            chart.Series.Add(series1);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            //chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSPieChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 39);
            DataPoint dataPoint2 = new DataPoint(0, 18);
            DataPoint dataPoint3 = new DataPoint(0, 15);
            DataPoint dataPoint4 = new DataPoint(0, 12);
            DataPoint dataPoint5 = new DataPoint(0, 5);


            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(64, 64, 64, 64);
            series1.ChartType = SeriesChartType.Pie;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(180, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);


            // Bind series data to chart
            chart.Series.Add(series1);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            //chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSLineChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Line;
            series1.BorderWidth = 3;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.MarkerSize = 8;
            series1.MarkerStyle = MarkerStyle.Circle;
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Line;
            series2.BorderWidth = 3;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.MarkerSize = 9;
            series2.MarkerStyle = MarkerStyle.Diamond;
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }
       
        public static Stream CreateMSRadarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Set series1 with width, maker, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Radar;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.MarkerBorderColor = Color.FromArgb(64, 64, 64);
            series1.MarkerSize = 9;
            series1.Name = "Series1";

            // Set series2 with width, maker, color, chart type and name           
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Radar;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.MarkerBorderColor = Color.FromArgb(64, 64, 64);
            series2.MarkerSize = 9;
            series2.Name = "Series2";

            // Bind series to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Populate series data
            double[] yValues = { 65.62, 75.54, 60.45, 34.73, 85.42, 55.9, 63.6, 55.1, 77.2 };
            double[] yValues2 = { 76.45, 23.78, 86.45, 30.76, 23.79, 35.67, 89.56, 67.45, 38.98 };
            string[] xValues = { "France", "Canada", "Germany", "USA", "Italy", "Spain", "Russia", "Sweden", "Japan" };
            chart.Series["Series1"].Points.DataBindXY(xValues, yValues);
            chart.Series["Series2"].Points.DataBindXY(xValues, yValues2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSPolarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Set series1 with width, maker, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.BorderWidth = 3;          
            series1.ChartType = SeriesChartType.Polar;           
            series1.MarkerBorderColor = Color.FromArgb(64, 64, 64);
            series1.MarkerColor = Color.LightSkyBlue;
            series1.MarkerSize = 7;
            series1.Name = "Series1";

            // Set series2 with width, maker, color, chart type and name           
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.BorderWidth = 3;
            series2.ChartType = SeriesChartType.Polar;
            series2.MarkerBorderColor = Color.FromArgb(64, 64, 64);
            series2.MarkerColor = Color.Gold;
            series2.MarkerSize = 7;
            series2.Name = "Series2";

            // Bind series to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Fill series data
            for (double angle = 0.0; angle <= 360.0; angle += 10.0)
            {
                double yValue = (1.0 + Math.Sin(angle / 180.0 * Math.PI)) * 10.0;
                chart.Series["Series1"].Points.AddXY(angle + 90.0, yValue);
                chart.Series["Series2"].Points.AddXY(angle + 270.0, yValue);
            }           

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSStacked100ColumnChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedColumn100;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedColumn100;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedColumn100;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            //chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSStacked100BarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedBar100;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedBar100;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedBar100;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSStacked100AreaChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedArea100;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedArea100;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedArea100;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            //chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSStackedAreaChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedArea;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedArea;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedArea;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

            
            return ms;
        }

        public static Stream CreateMSStackedBarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedBar;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedBar;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedBar;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            //chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSStackedColumnChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();
            Series series3 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);
            DataPoint dataPoint11 = new DataPoint(0, 3);
            DataPoint dataPoint12 = new DataPoint(0, 7);
            DataPoint dataPoint13 = new DataPoint(0, 2);
            DataPoint dataPoint14 = new DataPoint(0, 4);
            DataPoint dataPoint15 = new DataPoint(0, 8);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.StackedColumn;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.StackedColumn;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Set series3 with data, color, chart type and name
            series3.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series3.ChartType = SeriesChartType.StackedColumn;
            series3.Color = Color.FromArgb(220, 224, 64, 10);
            series3.Name = "Series3";
            series3.Points.Add(dataPoint11);
            series3.Points.Add(dataPoint12);
            series3.Points.Add(dataPoint13);
            series3.Points.Add(dataPoint14);
            series3.Points.Add(dataPoint15);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);
            chart.Series.Add(series3);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSColumnChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Column;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Column;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            //chartArea.Area3DStyle.Enable3D = false;
            chartArea.Area3DStyle.Enable3D = true;
            chartArea.Area3DStyle.Inclination = 15;
            chartArea.Area3DStyle.IsRightAngleAxes = false;
            chartArea.Area3DStyle.Perspective = 10;
            chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSBarChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Bar;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Bar;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            //chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }
        
        public static Stream CreateMSSplineAreaChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.SplineArea;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.SplineArea;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
            chartArea.BackColor = Color.WhiteSmoke;

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;
        }

        public static Stream CreateMSAreaChart()
        {
            // Create ms chart object
            Chart chart = new Chart();

            ChartArea chartArea = new ChartArea();
            Series series1 = new Series();
            Series series2 = new Series();

            // Input Data Point
            DataPoint dataPoint1 = new DataPoint(0, 5);
            DataPoint dataPoint2 = new DataPoint(0, 8.5);
            DataPoint dataPoint3 = new DataPoint(0, 9);
            DataPoint dataPoint4 = new DataPoint(0, 7);
            DataPoint dataPoint5 = new DataPoint(0, 5);
            DataPoint dataPoint6 = new DataPoint(0, 5);
            DataPoint dataPoint7 = new DataPoint(0, 4);
            DataPoint dataPoint8 = new DataPoint(0, 5);
            DataPoint dataPoint9 = new DataPoint(0, 3.5);
            DataPoint dataPoint10 = new DataPoint(0, 4);

            // Set series1 with data, color, chart type and name
            series1.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series1.ChartType = SeriesChartType.Area;
            series1.Color = Color.FromArgb(220, 65, 140, 240);
            series1.Name = "Series1";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            series1.Points.Add(dataPoint4);
            series1.Points.Add(dataPoint5);

            // Set series2 with data, color, chart type and name
            series2.BorderColor = Color.FromArgb(180, 26, 59, 105);
            series2.ChartType = SeriesChartType.Area;
            series2.Color = Color.FromArgb(220, 252, 180, 65);
            series2.Name = "Series2";
            series2.Points.Add(dataPoint6);
            series2.Points.Add(dataPoint7);
            series2.Points.Add(dataPoint8);
            series2.Points.Add(dataPoint9);
            series2.Points.Add(dataPoint10);

            // Bind series data to chart
            chart.Series.Add(series1);
            chart.Series.Add(series2);

            // Set chart 3D style
            chartArea.BackColor = Color.WhiteSmoke;
            chartArea.Area3DStyle.Enable3D = false;
            //chartArea.Area3DStyle.Enable3D = true;
            //chartArea.Area3DStyle.Inclination = 15;
            //chartArea.Area3DStyle.IsRightAngleAxes = false;
            //chartArea.Area3DStyle.Perspective = 10;
            //chartArea.Area3DStyle.Rotation = 10;
                      

            // Bind chart area to chart object
            chart.ChartAreas.Add(chartArea);

            // Modify chart size
            chart.Size = new Size(400, 300);

            // Render chart graphics to stream
            MemoryStream ms = new MemoryStream();
            chart.SaveImage(ms, ChartImageFormat.Png);

           
            return ms;  
        }
    }
}
