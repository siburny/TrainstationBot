using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using TrainStationBot.Properties;
using Rectangle = System.Drawing.Rectangle;
using Color = System.Drawing.Color;
using SystemColors = System.Drawing.SystemColors;
using System.IO;
using System.Text.RegularExpressions;

namespace TrainStationBot
{
    public partial class MainForm : Form
    {
        private static IntPtr GameWindowHandle;
        public static ScreenUtilities.RECT GameWindowPosition;

        private bool IsActive = false;
        private readonly Dictionary<string, List<Tuple<DateTime, int>>> Stats = new Dictionary<string, List<Tuple<DateTime, int>>>() { { "people", new List<Tuple<DateTime, int>>() }, { "mail", new List<Tuple<DateTime, int>>() } };
        private int Cycle = 0;

        public MainForm()
        {
            InitializeComponent();
            LoadImages();
        }

        private readonly List<Mat> ads = new List<Mat>();
        Mat arrow, box, whistle, levelup, bonus, balloon, dailyreward;

        private void LoadImages()
        {
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad1.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad2.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad3.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad4.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad5.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad6.png").CvtColor(ColorConversionCodes.BGR2GRAY));
            ads.Add(new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\ads\ad7.png").CvtColor(ColorConversionCodes.BGR2GRAY));

            arrow = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\arrow.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            box = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\box.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            whistle = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\whistle.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            levelup = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\levelup.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            bonus = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\bonus.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            balloon = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\balloon.png").CvtColor(ColorConversionCodes.BGR2GRAY);
            dailyreward = new Mat(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\dailyreward.png").CvtColor(ColorConversionCodes.BGR2GRAY);
        }

        private void StartStopButton_Click(object sender, EventArgs e)
        {
            if (IsActive)
            {
                IsActive = false;
                StartStopButton.Text = "START";
            }
            else
            {
                IsActive = true;
                StartStopButton.Text = "STOP";

                GameWindowHandle = ScreenUtilities.GetGameWindow();
                if (GameWindowHandle == default)
                {
                    MessageBox.Show("Can't find browser window");
                    Application.Exit();
                }

                ScreenUtilities.GetWindowRect(GameWindowHandle, ref GameWindowPosition);

                Thread child = new Thread(new ThreadStart(this.TimerLoop));
                child.Start();
            }
        }

        private Mat screen_rgb;
        private Mat screen_gray;
        private void TimerLoop()
        {
            while (true)
            {
                AddOutput("Analyzing screenshot");

                using (var image = ScreenUtilities.CaptureWindow(GameWindowHandle))
                {
                    screen_rgb = image.ToMat();
                    screen_gray = screen_rgb.CvtColor(ColorConversionCodes.BGRA2GRAY);

                    if (!IsActive) return;
                    if (Cycle % 10 == 0) CheckForAds();

                    if (!IsActive) return;
                    CheckForBalloon();

                    if (!IsActive) return;
                    CheckForLevelup();

                    if (!IsActive) return;
                    CheckForDailyReward();

                    if (!IsActive) return;
                    CheckForBox();

                    if (!IsActive) return;
                    CheckForBonus();

                    if (!IsActive) return;
                    CheckForTrains();

                    if (!IsActive) return;
                    CheckForWhistle();

                    if (!IsActive) return;
                    DoOCR();

                    screen_rgb.SaveImage(@"c:\Projects\C#\TrainStationBot\TrainStationBot\images\output\screen-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".jpg");

                    screen_rgb.Dispose();
                    screen_gray.Dispose();
                }

                Application.DoEvents();
                if (!IsActive) return;

                Cycle++;
            }
        }

        private void DoOCR()
        {
            try
            {
                var bitmap = screen_rgb.ToBitmap();
                string s_people = "",
                    s_mail = "";

                using (var engine = new Tesseract.TesseractEngine(@"./tessdata", "eng", Tesseract.EngineMode.Default))
                {
                    engine.SetVariable("tessedit_char_whitelist", "0123456789");
                    var people = bitmap.Clone(new Rectangle(230, 250, 130, 25), System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    using (var page = engine.Process(people))
                    {
                        s_people = Regex.Replace(page.GetText(), "[^0-9]", "");
                    }

                    var mail = bitmap.Clone(new Rectangle(395, 250, 182, 25), System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    using (var page = engine.Process(mail))
                    {
                        s_mail = Regex.Replace(page.GetText(), "[^0-9]", "");
                    }
                }

                if (string.IsNullOrEmpty(s_people))
                    s_people = "0";
                if (string.IsNullOrEmpty(s_mail))
                    s_mail = "0";

                Stats["people"].Add(new Tuple<DateTime, int>(DateTime.Now, int.Parse(s_people)));
                Stats["mail"].Add(new Tuple<DateTime, int>(DateTime.Now, int.Parse(s_mail)));

                File.AppendAllLines("output.json", new string[] { "{\"time\": \"" + DateTime.Now.ToString("o") + "\",\"people\": " + s_people + ", \"mail\": " + s_mail + " }" });

                if (!IsActive) return;
                Invoke(new Action(() => { PictureBoxOutput.Image = bitmap; }));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + e.StackTrace, "DoOCR Exception");
            }
        }

        private void CheckForAds()
        {
            foreach (var ad in ads)
            {
                Detect("ad", ad, out double maxValue, out Point maxLocation);

                if (maxValue > 0.9)
                {
                    AddOutput("Found AD at " + maxLocation.X + "x" + maxLocation.Y + ": clicked");
                    MouseOperations.MouseClick(maxLocation.X + ad.Width - 10, maxLocation.Y + 10);

                    UpdateScore(TextboxScore1, "AD: " + maxValue.ToString(), Color.LightGreen);
                }
                else
                {
                    UpdateScore(TextboxScore1, "AD: " + maxValue.ToString(), SystemColors.Control);
                }
            }
        }

        private void CheckForBox()
        {
            Detect("box", box, out double maxValue, out Point maxLocation);

            if (maxValue > 0.85)
            {
                AddOutput("Found BOX at " + maxLocation.X + "x" + maxLocation.Y + ": clicked it");
                MouseOperations.MouseClick(maxLocation.X + 20, 750);

                WaitNSeconds(100);

                MouseOperations.MouseClick(840, 812);

                UpdateScore(TextboxScore2, "BOX: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore2, "BOX: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForBalloon()
        {
            Detect("balloon", balloon, out double maxValue, out Point maxLocation);

            if (maxValue > 0.9)
            {
                AddOutput("Found BALLOON: " + maxLocation.X + "x" + maxLocation.Y);
                MouseOperations.MouseClick(maxLocation.X + 20, maxLocation.Y + 15);

                UpdateScore(TextboxScore3, "BALLOON: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore3, "BALLOON: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForLevelup()
        {
            Detect("LEVEL UP", levelup, out double maxValue, out Point maxLocation);

            if (maxValue > 0.9)
            {
                AddOutput("Found LEVEL UP: " + maxLocation.X + "x" + maxLocation.Y);
                MouseOperations.MouseClick(maxLocation.X + 20, 750);
                WaitNSeconds(100);

                MouseOperations.MouseClick(840 + 20, 884);

                UpdateScore(TextboxScore4, "LEVEL UP: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore4, "LEVEL UP: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForDailyReward()
        {
            Detect("reward", dailyreward, out double maxValue, out Point maxLocation);

            if (maxValue > 0.9)
            {
                AddOutput("Found DAILY REWARD: " + maxLocation.X + "x" + maxLocation.Y);
                MouseOperations.MouseClick(840 + 20, 779);

                UpdateScore(TextboxScore8, "DAILY REWARD: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore8, "DAILY REWARD: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForBonus()
        {
            Detect("bonus", bonus, out double maxValue, out Point maxLocation);

            if (maxValue > 0.9)
            {
                AddOutput("Found BONUS: " + maxLocation.X + "x" + maxLocation.Y);
                MouseOperations.MouseClick(maxLocation.X + 20, maxLocation.Y + 20);

                UpdateScore(TextboxScore5, "BONUS: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore5, "BONUS: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForTrains()
        {
            Detect("train", arrow, out double maxValue, out Point maxLocation);

            if (maxValue > 0.9)
            {
                AddOutput("Found TRAIN: " + maxLocation.X + "x" + maxLocation.Y);

                MouseOperations.MouseClick(900, 810);

                WaitNSeconds(250);

                MouseOperations.MouseClick(900, 810);

                WaitNSeconds(250);

                MouseOperations.MouseClick(500, 906);

                WaitNSeconds(250);

                MouseOperations.MouseClick(920, 730);

                UpdateScore(TextboxScore6, "TRAIN: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore6, "TRAIN: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void CheckForWhistle()
        {
            Detect("whistle", whistle, out double maxValue, out Point maxLocation);

            if (maxValue > 0.95)
            {
                AddOutput("Found WHISTLE: " + maxLocation.X + "x" + maxLocation.Y);
                MouseOperations.MouseClick(maxLocation.X + 20, maxLocation.Y + 20);

                UpdateScore(TextboxScore7, "WHISTLE: " + maxValue.ToString(), Color.LightGreen);
            }
            else
            {
                UpdateScore(TextboxScore7, "WHISTLE: " + maxValue.ToString(), SystemColors.Control);
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            RestoreWindowPosition();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            IsActive = false;
            SaveWindowPosition();
        }

        private void RestoreWindowPosition()
        {
            if (Settings.Default.HasSetDefaults)
            {
                this.WindowState = Settings.Default.WindowState;
                this.Location = Settings.Default.Location;
                this.Size = Settings.Default.Size;
            }
        }

        private void SaveWindowPosition()
        {
            Settings.Default.WindowState = this.WindowState;

            if (this.WindowState == FormWindowState.Normal)
            {
                Settings.Default.Location = this.Location;
                Settings.Default.Size = this.Size;
            }
            else
            {
                Settings.Default.Location = this.RestoreBounds.Location;
                Settings.Default.Size = this.RestoreBounds.Size;
            }

            Settings.Default.HasSetDefaults = true;

            Settings.Default.Save();
        }

        private void AddOutput(string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AddOutput(text)));
            }
            else
            {
                OutputListBox.Items.Add(DateTime.Now.ToString() + ": " + text);
                OutputListBox.TopIndex = OutputListBox.Items.Count - 1;
            }
        }

        private void UpdateScore(Control control, String text, System.Drawing.Color color)
        {
            if (control.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateScore(control, text, color)));
            }
            else
            {
                control.Text = text;
                control.BackColor = color;
            }
        }

        private void WaitNSeconds(int ms)
        {
            System.Threading.Thread.Sleep(ms);
            return;
        }

        private void Detect(string name, Mat template, out double maxValue, out Point maxLocation)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            using (Mat result = screen_gray.MatchTemplate(template, TemplateMatchModes.CCoeffNormed))
            {
                watch.Stop();
                Console.WriteLine("Time: " + 1.0 * watch.ElapsedMilliseconds / 1000);

                Cv2.MinMaxLoc(result, out _, out maxValue, out _, out maxLocation);
                Console.WriteLine("maxVal: " + maxValue + " maxLoc: " + maxLocation);

                if (maxValue > 0.9)
                {
                    screen_rgb.Rectangle(new Rect(maxLocation.X, maxLocation.Y, template.Width, template.Height), Scalar.LightGreen, 3);
                    screen_rgb.PutText(name, new Point(maxLocation.X, maxLocation.Y), HersheyFonts.HersheyPlain, 2, Scalar.White, 3);
                }
                else if (maxValue > 0.5)
                {
                    screen_rgb.Rectangle(new Rect(maxLocation.X, maxLocation.Y, template.Width, template.Height), Scalar.LightCoral, 3);
                    screen_rgb.PutText(name, new Point(maxLocation.X, maxLocation.Y), HersheyFonts.HersheyPlain, 2, Scalar.White, 3);
                }
            }
        }
    }
}