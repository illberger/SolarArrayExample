using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static SPACalculator.SPACalculator;

namespace SPACalculator
{
    /// <summary>
    /// Windows forms implementation to calculate solar panel output in central gothenburg based on https://github.com/xeqlol/SolarPositionAlgorithm
    /// Neglects temperature, shading obstructions, (assumes overall efficiency of 85% and panel efficiency of 15-25)
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            ConfigureChart(chart1, "Solar Array Output", "Time of day", "Power Output (Watts)");
            ConfigureChart(chart2, "Simulated Panel Output", "Time of day", "Power Output (Watts)");

            // Efficiency and panel setup
            double overallEfficiency = 0.85; // Inverter/system efficiency
            double panelEfficiency = 0.2059; // Panel conversion efficiency
            double panelArea = 1.755 * 1.038; // Panel area (m²)
            int nwwPanelCount = 18; // Number of panels in NWW string
            int seePanelCount = 26; // Number of panels in SEE string
            double nominalPower = 375; // Nominal power per panel (W)
            double panelSlope = 30; // Tilt angle of roof (degrees)
            double azimuthNW = 285; // Azimuth of roof with 18 solar panels relative to true north
            double azimuthSE = 105; // Azimuth of roof with 26 solar panels relative to true north
            double latitude = 57.708870; // Central Gothenburg latitude
            double longitute = 11.974560; // Central Gothenburg longitute

            // Prognosticated solar radiance data Gothenburg W/m^2 fetched from a third party 2024-12-27.
            Dictionary<int, double> radiationData = new Dictionary<int, double>
            {
                {9, 1}, {10, 51}, {11, 111}, {12, 141}, {13, 134}, {14, 92}, {15, 28}
            };

            AddChartSeries(chart1, "NWW String", SeriesChartType.Line, "Blue");
            AddChartSeries(chart1, "SEE String", SeriesChartType.Line, "Orange");
            AddChartSeries(chart2, "NWW Roof", SeriesChartType.Line, "Blue");
            AddChartSeries(chart2, "SEE Roof", SeriesChartType.Line, "Orange");

            double? sunriseAzimuth = null;
            double? sunsetAzimuth = null;
            DateTime? sunriseTime = null;
            DateTime? sunsetTime = null;
            bool sunHasRisen = false;

            for (int hour = 0; hour < 24; hour++)
            {
                // Date for observers orientation to sun.
                var dateTime = new DateTime(2025, 06, 15, hour, 0, 0);

                SolarPosition solarPosition = GetSolarPosition(dateTime, latitude, longitute, +1.0);

                double radiation = radiationData.ContainsKey(hour) ? radiationData[hour] : 0;

                if (!sunHasRisen && solarPosition.Altitude > 0)
                {
                    sunriseAzimuth = solarPosition.Azimuth;
                    sunriseTime = dateTime;
                    sunHasRisen = true;
                }

                if (sunHasRisen && solarPosition.Altitude < 0)
                {
                    sunsetAzimuth = solarPosition.Azimuth;
                    sunsetTime = dateTime;
                    sunHasRisen = false;
                }

                /* Debug statement for suns altitude at DateAndTime, can cross-reference with https://www.suncalc.org/
                //MessageBox.Show($"{solarPosition.Altitude}, hour: {hour}");
                */

                double nwwOrientationFactor = CalculateOrientationFactor(solarPosition.Azimuth, azimuthNW, solarPosition.Zenith, panelSlope); // NW roof
                double seeOrientationFactor = CalculateOrientationFactor(solarPosition.Azimuth, azimuthSE, solarPosition.Zenith, panelSlope); // SE roof

                double nwwPower = CalculateStringAdjustedNominalPower(nominalPower, nwwPanelCount, panelEfficiency, overallEfficiency, nwwOrientationFactor);
                double seePower = CalculateStringAdjustedNominalPower(nominalPower, seePanelCount, panelEfficiency, overallEfficiency, seeOrientationFactor);


                double nwwStringPower = CalculateStringPower(radiation, solarPosition, azimuthNW, panelSlope, nwwPanelCount, panelArea, panelEfficiency, overallEfficiency);
                double seeStringPower = CalculateStringPower(radiation, solarPosition, azimuthSE, panelSlope, seePanelCount, panelArea, panelEfficiency, overallEfficiency);

                chart1.Series["NWW String"].Points.AddXY(hour, nwwStringPower);
                chart1.Series["SEE String"].Points.AddXY(hour, seeStringPower);
                chart2.Series["NWW Roof"].Points.AddXY(hour, nwwPower);
                chart2.Series["SEE Roof"].Points.AddXY(hour, seePower);
            }

            if (sunriseTime.HasValue && sunsetTime.HasValue)
            {
                MessageBox.Show($"Sunrise at {sunriseTime.Value:HH:mm} (Azimuth: {sunriseAzimuth:F2}°)\n" +
                                $"Sunset at {sunsetTime.Value:HH:mm} (Azimuth: {sunsetAzimuth:F2}°)");
            }
            else
            {
                MessageBox.Show("The sun does not rise or set on this day at the given location."); // Above polar circle.
            }
        }


        /// <summary>
        /// Helper method to configure chart object.
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="chartAreaName"></param>
        /// <param name="xAxisTitle"></param>
        /// <param name="yAxisTitle"></param>
        private static void ConfigureChart(Chart chart, string chartAreaName, string xAxisTitle, string yAxisTitle)
        {
            chart.Series.Clear();
            chart.ChartAreas.Clear();
            chart.Legends.Clear();

            ChartArea chartArea = new ChartArea(chartAreaName)
            {
                AxisX = { Title = xAxisTitle },
                AxisY = { Title = yAxisTitle }
            };
            chart.ChartAreas.Add(chartArea);

            Legend legend = new Legend
            {
                Name = "Legend",
                Font = new System.Drawing.Font("Arial", 12),
                Docking = Docking.Top
            };
            chart.Legends.Add(legend);
        }

        
        /// <summary>
        /// Helper method to add series for plotting to chart object.
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="name"></param>
        /// <param name="chartType"></param>
        /// <param name="colorName"></param>
        private static void AddChartSeries(Chart chart, string name, SeriesChartType chartType, string colorName)
        {
            Series series = new Series(name)
            {
                ChartType = chartType,
                Legend = "Legend",
                IsVisibleInLegend = true,
                Color = System.Drawing.Color.FromName(colorName),
                BorderWidth = 3
            };
            chart.Series.Add(series);
        }

        /// <summary>
        /// Method to adjust Pnom of solar string/array/panel(s) based on actual solar position.
        /// </summary>
        /// <param name="nominalPower"></param>
        /// <param name="panelCount"></param>
        /// <param name="panelEfficiency"></param>
        /// <param name="overallEfficiency"></param>
        /// <param name="orientationFactor"></param>
        /// <returns></returns>
        private double CalculateStringAdjustedNominalPower(double nominalPower, int panelCount, double panelEfficiency, double overallEfficiency, double orientationFactor)
        {
            double totalNominalPower = nominalPower * panelCount;
            double nSystem = orientationFactor * overallEfficiency;
            return totalNominalPower * nSystem;
        }


        /// <summary>
        /// Method to derive an adjustment factor for the panels solar position.
        /// </summary>
        /// <param name="sunAzimuth"></param>
        /// <param name="roofAzimuth">Azimuth/orientation of surface perpendicular to solar panel/string/array relative to true north</param>
        /// <param name="sunZenith"></param>
        /// <param name="panelSlope">Vertical tilt/Tilt Angle of Panel and/or surface perpendicular to panel/solar string/array</param>
        /// <returns></returns>
        private double CalculateOrientationFactor(double sunAzimuth, double roofAzimuth, double sunZenith, double panelSlope)
        {
            // If the sun is below the horizon (zenith > 90) its out of sight 
            if (sunZenith > 90)
                return 0;

            double sunZenithRad = DegreeToRadian(sunZenith);
            double panelSlopeRad = DegreeToRadian(panelSlope);
            double azimuthDifferenceRad = DegreeToRadian(sunAzimuth - roofAzimuth);

            double cosTheta = Math.Cos(sunZenithRad) * Math.Cos(panelSlopeRad) +
                              Math.Sin(sunZenithRad) * Math.Sin(panelSlopeRad) * Math.Cos(azimuthDifferenceRad);

            // Return the cosine factor if positive, otherwise 0 (sun is behind the panels)
            return (cosTheta > 0) ? cosTheta : 0;
        }



        // Calculate effective radiation for a panel orientation
        private double CalculateEffectiveRadiation(double radiation, SolarPosition solarPosition, double panelAzimuth, double panelSlope)
        {
            // If the sun is below the horizon (zenith > 90°)
            if (solarPosition.Zenith > 90)
            {
                return 0;
            }

            double solarZenithRad = DegreeToRadian(solarPosition.Zenith); // Sun zenith angle
            double solarAzimuthRad = DegreeToRadian(solarPosition.Azimuth);
            double panelAzimuthRad = DegreeToRadian(panelAzimuth);
            double panelSlopeRad = DegreeToRadian(panelSlope); // Panel tilt (slope)

            // Calculate cosine of the angle of incidence
            double cosTheta = Math.Cos(solarZenithRad) * Math.Cos(panelSlopeRad) +
                              Math.Sin(solarZenithRad) * Math.Sin(panelSlopeRad) * Math.Cos(solarAzimuthRad - panelAzimuthRad);

            // Return effective radiation (only positive values are valid)
            return (cosTheta > 0) ? radiation * cosTheta : 0;
        }



        /// <summary>
        /// Calculate Power of string taking into account actual radiance W/m^2.
        /// </summary>
        /// <param name="radiation"></param>
        /// <param name="solarPosition"></param>
        /// <param name="panelAzimuth"></param>
        /// <param name="panelSlope"></param>
        /// <param name="panelCount"></param>
        /// <param name="panelArea"></param>
        /// <param name="panelEfficiency"></param>
        /// <param name="overallEfficiency"></param>
        /// <returns></returns>
        private double CalculateStringPower(double radiation, SolarPosition solarPosition, double panelAzimuth, double panelSlope, int panelCount, double panelArea, double panelEfficiency, double overallEfficiency)
        {
            // Calculate effective radiation based on panel orientation and tilt
            double effectiveRadiation = CalculateEffectiveRadiation(radiation, solarPosition, panelAzimuth, panelSlope);

            // Calculate power output
            return (effectiveRadiation * panelArea * panelCount) * panelEfficiency * overallEfficiency;
        }


        
        private double DegreeToRadian(double degree)
        {
            return degree * Math.PI / 180.0;
        }

        /// <summary>
        /// Get the Solar position relative to an observer of latitute, longitute and DateAndTime.
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="latitude"></param>
        /// <param name="longitude"></param>
        /// <param name="timezoneOffset"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public SolarPosition GetSolarPosition(DateTime dateTime, double latitude, double longitude, double timezoneOffset)
        {
            var spa = new SPAData
            {
                Year = dateTime.Year,
                Month = dateTime.Month,
                Day = dateTime.Day,
                Hour = dateTime.Hour,
                Minute = dateTime.Minute,
                Second = dateTime.Second,
                Timezone = timezoneOffset,
                Longitude = longitude,
                Latitude = latitude,
                Elevation = 20,
                Pressure = 1013,
                Temperature = 5,
                Slope = 0,
                AzmRotation = 150,
                AtmosRefract = 0.5667,
                Function = CalculationMode.SPA_ALL
            };

            var result = SPACalculator.SPACalculate(ref spa);
            if (result != 0)
            {
                throw new Exception($"SPA Calculation Error: {result}");
            }
            else
            {
                //MessageBox.Show($"Julian Day: {spa.Jd}, L: degrees: {spa.L}, B: {spa.B}, R: {spa.R}, H: {spa.H}, Zenith: {spa.Zenith}, Azimuth: {spa.Azimuth}, Altiude: {spa.Incidence: {spa.Incidence}");
            }

            return new SolarPosition
            {
                Azimuth = spa.Azimuth,
                Altitude = 90 - spa.Zenith,
                Zenith = spa.Zenith
            };
        }
    }

    /// <summary>
    /// Values extrapolated from SPA.
    /// </summary>
    public class SolarPosition
    {
        public double Azimuth { get; set; }
        public double Altitude { get; set; }
        public double Zenith { get; set; }
    }
}
