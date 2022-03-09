using CommonLib.EpplusErrorReport;
using CommonLib.Utility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using ProfileTool_PMIC.Output;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;
using ProfileTool_PMIC.Reader;
using Chart = System.Windows.Forms.DataVisualization.Charting.Chart;
using ChartArea = System.Windows.Forms.DataVisualization.Charting.ChartArea;
using Font = System.Drawing.Font;
using Legend = System.Windows.Forms.DataVisualization.Charting.Legend;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;

namespace ProfileTool_PMIC
{
    public class PowerpointImage
    {
        public Profile ProfileFile { get; set; }
        public string ImageFile { get; set; }
    }

    public class ProfileAnalysisMain
    {
        private string _profilePath;
        private string _outputPath;
        private string _tempPath;

        private ProfileToolForm _profileToolForm;
        private readonly string _outputFile;
        private int _imageCountPerSlide;
        private List<Profile> _profileFiles = new List<Profile>();
        private Dictionary<int, List<PowerpointImage>> _powerpointImageDic = new Dictionary<int, List<PowerpointImage>>();
        private int _loopCount;
        private const int PowerpointWidth = 1280;
        private const int PowerpointHeight = 720;
        private const int MultiPinHeight = 1000;
        private int _chartMergeCount;
        private const int MaxChartCount = 1700000;
        private const int DevideChartCount = 20000;
        private string _maxSummaryChartPath = null;
        private Presentation Presentation;

        public ProfileAnalysisMain(ProfileToolForm profileToolForm, string outputFile)
        {
            _profileToolForm = profileToolForm;
            _outputFile = outputFile;
            _profilePath = profileToolForm.FileOpen_ProfilePath1.ButtonTextBox.Text;
            _outputPath = profileToolForm.FileOpen_OutputPath2.ButtonTextBox.Text;
            _loopCount = int.Parse(_profileToolForm.textBoxLoopCount.Text);
            _tempPath = Path.Combine(_outputPath, "Temp");
            if (Directory.Exists(_tempPath))
                Directory.Delete(_tempPath, true);
            Directory.CreateDirectory(_tempPath);
        }

        public void WorkFlow()
        {
            EpplusErrorManager.ResetError();
            _profileFiles = new List<Profile>();
            _powerpointImageDic = new Dictionary<int, List<PowerpointImage>>();
            _profileToolForm.AppendText(string.Format("Reading Profile Files ..."), Color.Blue);
            var profileFiles = Directory.GetFiles(_profilePath, "*Profile-*.txt", SearchOption.AllDirectories);
            profileFiles = SortProfileFiles(profileFiles);

            bool flag = _profileToolForm.checkBox_Power.Checked;
            var profilePath2 = _profileToolForm.FileOpen_ProfilePath2.ButtonTextBox.Text;
            var profileFiles2 = !string.IsNullOrEmpty(profilePath2) ?
                Directory.GetFiles(profilePath2, "*Profile-*.txt", SearchOption.AllDirectories).ToList() : null;

            if (_profileToolForm.radioButtonIndividual.Checked)
            {
                ByIndividual(profileFiles, flag, profileFiles2);
            }
            else
            {
                for (int index = 0; index < profileFiles.Length; index++)
                {
                    var profilFile = profileFiles[index];
                    _profileToolForm.AppendText(string.Format("Reading file {0} - {1}...", Path.GetFileNameWithoutExtension(profilFile), index + 1 + "/" + profileFiles.Count()), Color.Blue);
                    _profileFiles.Add(new ProfileReader().ReadProfileFile(profilFile, flag, profileFiles2));
                }
                _profileFiles = _profileFiles.OrderBy(x => x.Date).ThenBy(x => x.Pin).ThenBy(x => x.Site).ToList();

                _powerpointImageDic = new Dictionary<int, List<PowerpointImage>>();
                if (_profileToolForm.radioButtonMerge.Checked)
                    _powerpointImageDic = GenImageMerge(_profileFiles);
            }

            if (_profileToolForm.checkBox_MultiPins.Checked)
            {
                int cnt = _powerpointImageDic.Last().Key + 1;
                foreach (var powerpointImageDic in GenChartByMultiPins(_profileFiles))
                {
                    _powerpointImageDic.Add(cnt, powerpointImageDic.Value);
                    cnt++;
                }
            }

            GenPowerpoint();

            GenErrorReoprt();

            GenSummaryReport();

            GenMaxSummaryByPinReport();

            GenMaxSummaryReport();

            GenMaxChartToPPT();

            _profileToolForm.AppendText(string.Format("The output excel file is {0} ...", _outputFile + ".xlsx"), Color.Blue);
            _profileToolForm.AppendText(string.Format("The output powerpoint file is {0} ...", _outputFile + ".pptx"), Color.Blue);
        }

        private void GenChartByMax(ExcelWorksheet sheet, int rowEnd, int colStart, int colEnd)
        {
            List<string> itemName = new List<string>();
            List<double> pinValue = new List<double>();
            for (int i = 2; i <= rowEnd; i++)
            {
                itemName.Add(sheet.Cells[i, 1].Value.ToString());
                for (int j = colStart; j <= colEnd; j++)
                {
                    pinValue.Add(Convert.ToDouble(sheet.Cells[i, j].Value));
                }
            }

            using (Chart chart1 = new Chart())
            {
                chart1.Height = PowerpointHeight;
                chart1.Width = PowerpointWidth + 1000;
                ChartArea area = chart1.ChartAreas.Add("chartArea");
                area.BorderDashStyle = ChartDashStyle.Solid;
                area.AxisX.Title = "Items";
                area.AxisX.LabelStyle.Angle = -90;
                area.AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount;
                area.AxisX.IsLabelAutoFit = false;

                area.AxisX.TitleFont = new Font("Arial", 28, FontStyle.Regular);
                area.AxisY.MajorTickMark.TickMarkStyle = TickMarkStyle.None;
                area.AxisY.Title = "Max value";
                area.AxisY.TitleFont = new Font("Arial", 28, FontStyle.Regular);

                area.Position.X = 10;
                area.Position.Y = 15;
                area.Position.Height = 84;
                area.Position.Width = 80;

                Legend legend = new Legend();
                legend.Docking = Docking.Top;
                legend.Alignment = StringAlignment.Center;
                legend.Position.Auto = true;
                legend.BackColor = Color.Transparent;
                chart1.Legends.Add(legend);

                var pins = _profileFiles.Select(x => x.Pin).Distinct().ToList();
                foreach (var pin in pins)
                {
                    Series series = new Series(pin);
                    series.ChartType = SeriesChartType.StackedColumn;
                    chart1.Series.Add(series);
                }

                for (int j = colStart; j <= colEnd; j++)
                {
                    List<double> xValue = new List<double>();
                    for (int rowidx = 2; rowidx <= rowEnd; rowidx++)
                    {
                        xValue.Add(Convert.ToDouble(sheet.Cells[rowidx, j].Value));
                    }
                    chart1.Series[j - colStart].Points.DataBindXY(itemName, xValue);

                }

                area.RecalculateAxesScale();
                var imagePath = _tempPath + "\\maxSummaryChart.png";
                _profileToolForm.AppendText(string.Format("Exporting file {0} ...", Path.GetFileNameWithoutExtension(imagePath)), Color.Blue);
                chart1.SaveImage(imagePath, ChartImageFormat.Png);
                _maxSummaryChartPath = imagePath;
            }

        }

        private Dictionary<int, List<PowerpointImage>> GenChartByMultiPins(List<Profile> profileFiles)
        {
            Dictionary<int, List<PowerpointImage>> dic = new Dictionary<int, List<PowerpointImage>>();
            var pinGroups = profileFiles.GroupBy(x => x.Item + x.Date).ToList();
            int groupCnt = 0;
            foreach (var pinGroup in pinGroups)
            {
                var siteGroups = pinGroup.GroupBy(x => x.Site);
                List<PowerpointImage> images = new List<PowerpointImage>();
                foreach (var siteGroup in siteGroups)
                {
                    var sortSiteGroup = siteGroup.ToList().OrderBy(x => x.Date);
                    var site = sortSiteGroup.First().Site;
                    var item = sortSiteGroup.First().Item;
                    var name = "All_Pins" + "_Site" + site + "_" + item;
                    var currentChartCount = sortSiteGroup.Sum(x => x.Value.Count);
                    using (Chart chart = new Chart())
                    {
                        double startTime = 0;
                        #region add chart Series
                        if (currentChartCount < MaxChartCount)
                        {
                            for (int i = 0; i < sortSiteGroup.Count(); i++)
                            {
                                var row = sortSiteGroup.ElementAt(i);
                                int dateCnt = 0;
                                Series series = new Series(row.Item + i);
                                series.ChartType = SeriesChartType.Line;
                                series.LegendText = row.Pin;
                                chart.Series.Add(series);
                                List<double> xAxis = new List<double>();
                                List<double> yAxis = row.Value;
                                foreach (var value in row.Value)
                                {
                                    xAxis.Add(startTime + dateCnt / row.SampleRate);
                                    dateCnt++;
                                }
                                series.Points.DataBindXY(xAxis, yAxis);
                            }
                        }
                        else
                        {
                            _chartMergeCount = currentChartCount / DevideChartCount;
                            _chartMergeCount = _chartMergeCount == 0 ? 1 : _chartMergeCount;
                            for (int i = 0; i < sortSiteGroup.Count(); i++)
                            {
                                var row = sortSiteGroup.ElementAt(i);
                                chart.Series.Add(MergeSeries(row, row.Item + i, ref startTime, false));
                            }
                        }

                        #endregion

                        SetChartFormat(chart, name);
                        var imageFile = _tempPath + "\\" + name + ".png";
                        _profileToolForm.AppendText(string.Format("Exporting file {0} ...", Path.GetFileNameWithoutExtension(imageFile)), Color.Blue);
                        chart.SaveImage(imageFile, ImageFormat.Png);
                        chart.Dispose();
                        PowerpointImage powerpointImage = new PowerpointImage();
                        powerpointImage.ImageFile = imageFile;
                        images.Add(powerpointImage);
                    }
                }
                groupCnt++;
                dic.Add(groupCnt, images);
            }
            return dic;
        }

        private static string[] SortProfileFiles(string[] profileFiles)
        {

            Dictionary<string, Profile> dic = new Dictionary<string, Profile>();
            ProfileReader profileReader = new ProfileReader();
            foreach (var profileFile in profileFiles)
            {
                var Profile = profileReader.ReadprofileFileWithoutValue(profileFile);
                dic.Add(profileFile, Profile);
            }
            profileFiles = dic.OrderBy(x => x.Value.Pin).ThenBy(x => x.Value.Site).ThenBy(x => x.Value.Date).Select(x => x.Key).ToArray();
            return profileFiles;
        }

        private void GenErrorReoprt()
        {
            using (ExcelPackage errorReport = new ExcelPackage(new FileInfo(Path.Combine(_outputPath, "Error.xlsx"))))
            {
                EpplusErrorManager.GenErrorReport(errorReport, null);
                if (errorReport.Workbook.Worksheets.Count > 0)
                    errorReport.Save();
            }
        }

        private void GenSummaryReport()
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(Path.Combine(Path.ChangeExtension(_outputFile, ".xlsx")))))
            {
                var images = _powerpointImageDic.Select(x => x.Value).SelectMany(x => x);
                foreach (var _profileFile in _profileFiles)
                {
                    if (images.Where(x => x.ProfileFile != null).Any(x => x.ProfileFile.FilePath.Equals(_profileFile.FilePath, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var row = images.First(x => x.ProfileFile.FilePath.Equals(_profileFile.FilePath, StringComparison.CurrentCultureIgnoreCase));
                        _profileFile.HyperLink = row.ImageFile;
                    }
                }

                var wroksheet = excel.Workbook.AddSheet("Summary");
                wroksheet.Cells[1, 1].LoadFromCollection(_profileFiles, true);
                wroksheet.SetFormula(3);

                ExcelAddress foramatRangeAddress1 = new ExcelAddress("I:I");
                var condition1 = wroksheet.ConditionalFormatting.AddExpression(foramatRangeAddress1);
                condition1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                condition1.Style.Fill.BackgroundColor.Color = Color.Red;
                condition1.Formula = string.Format("IF(AND(ISNUMBER(I1),I1<>L1),1,0)");

                ExcelAddress foramatRangeAddress2 = new ExcelAddress("J:J");
                var condition2 = wroksheet.ConditionalFormatting.AddExpression(foramatRangeAddress2);
                condition2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                condition2.Style.Fill.BackgroundColor.Color = Color.Red;
                condition2.Formula = string.Format("IF(AND(ISNUMBER(J1),J1<>M1),1,0)");

                ExcelAddress foramatRangeAddress3 = new ExcelAddress("K:K");
                var condition3 = wroksheet.ConditionalFormatting.AddExpression(foramatRangeAddress3);
                condition3.Style.Fill.PatternType = ExcelFillStyle.Solid;
                condition3.Style.Fill.BackgroundColor.Color = Color.Red;
                condition3.Formula = string.Format("IF(AND(ISNUMBER(K1),K1<>N1),1,0)");
                excel.Save();
            }
        }

        private void GenMaxSummaryByPinReport()
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(Path.Combine(Path.ChangeExtension(_outputFile, ".xlsx")))))
            {
                var wroksheet = excel.Workbook.AddSheet("MaxSummaryByPin");
                var pins = _profileFiles.Select(x => x.Pin).Distinct().ToList();

                var headers = new List<string> { "", "MaxValue", "Item", "ChartType" };
                headers.AddRange(pins);
                headers.Add("Total");
                int cnt = 1;
                foreach (var header in headers)
                {
                    wroksheet.Cells[1, cnt].Value = header;
                    cnt++;
                }

                int row = 2;
                const int colStart = 5;
                foreach (var pinRow in pins)
                {
                    List<double> values = new List<double>();
                    wroksheet.Cells[row, 1].Value = pinRow;
                    double max = _profileFiles.Where(x => x.Pin.Equals(pinRow, StringComparison.CurrentCultureIgnoreCase)).Max(x => x.MaxAfterFilter);
                    var profileFile = _profileFiles.Where(x => x.Pin.Equals(pinRow, StringComparison.CurrentCultureIgnoreCase)).First(x => x.MaxAfterFilter == max);
                    wroksheet.Cells[row, 2].Value = string.Format("{0:F4}", max);
                    wroksheet.Cells[row, 3].Value = profileFile.Item;
                    wroksheet.Cells[row, 4].Value = profileFile.ChartType;
                    int col = colStart;
                    foreach (var pinCol in pins)
                    {
                        var profileFile2 = _profileFiles.Find(x => x.Site == profileFile.Site && x.Item == profileFile.Item &&
                            x.Pin.Equals(pinCol, StringComparison.CurrentCultureIgnoreCase) && x.ChartType == profileFile.ChartType && x.Date == profileFile.Date);

                        if (profileFile2 != null)
                        {
                            int index = (int)(profileFile.MaxIndex / profileFile.SampleRate * profileFile2.SampleRate);
                            if (index >= profileFile2.Value.Count)
                                index = profileFile2.Value.Count - 1;
                            if (pinRow == pinCol)
                                index = profileFile.MaxIndex;

                            if (string.IsNullOrEmpty(profileFile2.HyperLink))
                                wroksheet.Cells[row, col].Value = profileFile2.Value[index];
                            else
                            {
                                wroksheet.Cells[row, col].Value = "=HYPERLINK(\"" + profileFile2.HyperLink + "\",\"" + string.Format("{0:F4}", profileFile2.Value[index]) + "\")";
                                wroksheet.Cells[row, col].SetHyperLinkFormat();
                            }
                            values.Add(profileFile2.Value[index]);
                        }
                        col++;
                    }
                    wroksheet.Cells[row, col].Value = values.Sum();
                    row++;
                }

                wroksheet.Column(1).AutoFit();
                wroksheet.Column(2).AutoFit();
                wroksheet.Column(3).AutoFit();
                excel.Save();
            }
        }

        private void GenMaxSummaryReport()
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(Path.Combine(Path.ChangeExtension(_outputFile, ".xlsx")))))
            {
                var wroksheet = excel.Workbook.AddSheet("MaxSummary");
                var pins = _profileFiles.Select(x => x.Pin).Distinct().ToList();

                var headers = new List<string> { "Item", "FilePath", "ChartType", "Link" };
                headers.AddRange(pins);
                headers.Add("Total");
                int cnt = 1;
                foreach (var header in headers)
                {
                    wroksheet.Cells[1, cnt].Value = header;
                    cnt++;
                }

                int row = 2;

                var group = _profileFiles.GroupBy(x => x.Pin).First();
                double maxTotal = 0;
                int maxIndex = 0;
                const int colStart = 5;
                foreach (var profileFile in group)
                {
                    List<Profile> profiles = new List<Profile>();
                    foreach (var pinCol in pins)
                    {
                        var profileFile2 = _profileFiles.Find(x => x.Site == profileFile.Site && x.Item == profileFile.Item &&
                                                    x.Pin.Equals(pinCol, StringComparison.CurrentCultureIgnoreCase) && x.ChartType == profileFile.ChartType && x.Date == profileFile.Date);
                        if (profileFile2 != null)
                            profiles.Add(profileFile2);
                        else
                            _profileToolForm.AppendText(string.Format("Missing profile file for {0} - {1}...", pinCol, profileFile.FilePath), Color.Red);
                    }

                    if (profiles.Count() != pins.Count())
                        continue;

                    for (int i = 0; i < profileFile.Value.Count; i++)
                    {
                        List<double> maxList = new List<double>();
                        foreach (var profile in profiles)
                        {
                            int index = (int)(i / profileFile.SampleRate * profile.SampleRate);
                            if (index >= profile.Value.Count)
                                index = profile.Value.Count - 1;
                            maxList.Add(profile.Value[index]);
                        }

                        var total = maxList.Sum();
                        if (maxList.Count() == pins.Count())
                        {
                            if (total > maxTotal)
                            {
                                maxTotal = total;
                                maxIndex = i;
                            }
                        }
                    }

                    if (maxTotal != 0)
                    {
                        wroksheet.Cells[row, 1].Value = profileFile.Item;
                        wroksheet.Cells[row, 2].Value = profileFile.FilePath;
                        var site = profileFile.Site;
                        var item = profileFile.Item;
                        var date = profileFile.Item;
                        var name = "All_Pins" + "_Site" + site + "_" + item + "_" + date;
                        var hyperLink = _tempPath + "\\" + name + ".png";
                        wroksheet.Cells[row, 3].Value = "=HYPERLINK(\"" + hyperLink + "\",\"Link\")";
                        wroksheet.Cells[row, 3].SetHyperLinkFormat();

                        wroksheet.Cells[row, 4].Value = profileFile.ChartType;
                        int column = colStart;
                        List<double> totalValues = new List<double>();
                        foreach (var pinCol in pins)
                        {
                            var profileFile2 = _profileFiles.Find(x => x.Site == profileFile.Site && x.Item == profileFile.Item &&
                                x.Pin.Equals(pinCol, StringComparison.CurrentCultureIgnoreCase) && x.ChartType == profileFile.ChartType && x.Date == profileFile.Date);
                            if (profileFile2 != null)
                            {
                                int index = (int)(maxIndex / profileFile.SampleRate * profileFile2.SampleRate);
                                if (index >= profileFile2.Value.Count)
                                    index = profileFile2.Value.Count - 1;
                                wroksheet.Cells[row, column].Value = profileFile2.Value[index];
                                totalValues.Add(profileFile2.Value[index]);
                            }
                            column++;
                        }
                        wroksheet.Cells[row, column].Value = totalValues.Sum();
                        row++;
                    }

                }

                wroksheet.Column(1).AutoFit();
                wroksheet.Column(2).AutoFit();
                wroksheet.Column(3).AutoFit();

                int rowEnd = wroksheet.Dimension.End.Row;
                int colEnd = wroksheet.Dimension.End.Column - 1;

                GenChartByMax(wroksheet, rowEnd, colStart, colEnd);
                var chartsheet = excel.Workbook.AddSheet("MaxSummary_Chart");
                ExcelChart chart = chartsheet.Drawings.AddChart("chart", eChartType.ColumnStacked);
                ExcelChartSerie serie = null;

                chart.SetPosition(0, 0);
                chart.SetSize(1400, 500);

                chart.XAxis.Title.Text = "Items";
                chart.YAxis.Title.Text = "Max value";

                for (int j = colStart; j <= colEnd; j++)
                {
                    serie = chart.Series.Add(wroksheet.Cells[2, j, rowEnd, j], wroksheet.Cells[2, 1, rowEnd, 1]);
                    serie.HeaderAddress = wroksheet.Cells[1, j];
                }

                chart.AdjustPositionAndSize();
                excel.Workbook.Worksheets.MoveAfter("MaxSummary_Chart", "MaxSummary");
                excel.Save();
            }
        }

        private void GenPowerpoint()
        {
            if (_profileToolForm.radioButtonAuto.Checked)
            {
                if (_profileToolForm.radioButtonIndividual.Checked)
                {
                    if (_loopCount != 0 && _profileToolForm.checkBoxOnlyLast.Checked)
                        _imageCountPerSlide = 2;
                    else
                        _imageCountPerSlide = 1 + _loopCount;
                }
                else if (_profileToolForm.radioButtonMerge.Checked)
                    _imageCountPerSlide = _profileFiles.Select(x => x.Site).Distinct().Count();
                else
                    _imageCountPerSlide = _profileFiles.Select(x => x.Pin).Distinct().Count();
            }
            else
                _imageCountPerSlide = int.Parse(_profileToolForm.textBoxChartCount.Text);
            GenPowerPoint(_powerpointImageDic);
        }

        private void GenMaxChartToPPT()
        {
            Slides slides = Presentation.Slides;
            CustomLayout customLayoutMax = Presentation.SlideMaster.CustomLayouts[7];
            Slide slideMax = slides.AddSlide(2, customLayoutMax);
            int height2 = 350;
            int top2 = 140;
            slideMax.Shapes.AddPicture(_maxSummaryChartPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, top2, 700, height2);
            Presentation.SaveAs(_outputFile + ".pptx");
        }

        private void GenPowerPoint(Dictionary<int, List<PowerpointImage>> powerpointImageDic)
        {
            const int gap = 20;
            _profileToolForm.AppendText(string.Format("Generating power point ..."), Color.Blue);
            Application app = new Application();
            Presentation = app.Presentations.Add(MsoTriState.msoCTrue);
            Slides slides = Presentation.Slides;
            slides.AddSlide(1, Presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutTitle]);
            foreach (var imageFiles in powerpointImageDic)
            {
                int remainslide = imageFiles.Value.Count % _imageCountPerSlide;
                int slideLoop = remainslide == 0 ? imageFiles.Value.Count / _imageCountPerSlide : imageFiles.Value.Count / _imageCountPerSlide + 1;
                int cnt = 0;
                for (int i = 0; i < slideLoop; i++)
                {
                    if (_imageCountPerSlide == 1)
                    {
                        CustomLayout customLayout = Presentation.SlideMaster.CustomLayouts[7];
                        Slide slide = slides.AddSlide(slides.Count + 1, customLayout);
                        int height = 350;
                        int top = 140;
                        slide.Shapes.AddPicture(imageFiles.Value[cnt].ImageFile, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, top, 700, height);
                        if (imageFiles.Value[cnt].ProfileFile != null)
                        {
                            var testbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, top + height, 700, height);
                            testbox.TextFrame.TextRange.Text = imageFiles.Value[cnt].ProfileFile.FilePath;
                            testbox.TextFrame.TextRange.Font.Size = 8;
                            testbox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255);
                            testbox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                        }
                    }
                    else
                    {
                        CustomLayout customLayout = Presentation.SlideMaster.CustomLayouts[7];
                        Slide slide = slides.AddSlide(slides.Count + 1, customLayout);
                        for (int j = 0; j < _imageCountPerSlide; j++)
                        {
                            int remain = cnt % _imageCountPerSlide;
                            int height = (500 - (_imageCountPerSlide - 1) * gap) / _imageCountPerSlide;
                            int top = 10 + remain * (height + gap);
                            if (cnt < imageFiles.Value.Count)
                            {
                                slide.Shapes.AddPicture(imageFiles.Value[cnt].ImageFile, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, top, 700, height);
                                if (imageFiles.Value[cnt].ProfileFile != null)
                                {
                                    var testbox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, top + height, 700, height);
                                    testbox.TextFrame.TextRange.Text = imageFiles.Value[cnt].ProfileFile.FilePath;
                                    testbox.TextFrame.TextRange.Font.Size = 8;
                                    testbox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                                    testbox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255);
                                }
                            }
                            cnt++;
                        }
                    }
                }
            }
            var templateFile = Path.Combine(Directory.GetCurrentDirectory(), @"Config\Template.potx");
            Presentation.Slides[1].ApplyTheme(templateFile);
            Presentation.SaveAs(_outputFile + ".pptx");

            //presentation.Close();
            //app.Quit();
        }

        private int RGB(int r, int g, int b)
        {
            return r * 65536 + g * 256 + b;
        }

        private void ByIndividual(string[] profileFiles, bool byPower, List<string> profileFiles2)
        {
            ProfileReader profileReader = new ProfileReader();
            _profileFiles = new List<Profile>();

            var pulseWidth = double.Parse(_profileToolForm.textBoxPulseWidth.Text);
            var stdevSpec = int.Parse(_profileToolForm.textBoxStdev.Text);
            _profileToolForm.ToolStripProgressBar.ProgressBar.Maximum = profileFiles.Count();
            for (int index = 0; index < profileFiles.Length; index++)
            {
                var profileFile = profileFiles[index];
                List<PowerpointImage> powerpointImage = new List<PowerpointImage>();
                _profileToolForm.AppendText(string.Format("Parsing  file {0} - {1}...", Path.GetFileNameWithoutExtension(profileFile), index + 1 + "/" + profileFiles.Count()), Color.Blue);
                _profileToolForm.ProgressBarIncrement();

                var profileItem = profileReader.ReadProfileFile(profileFile, byPower, profileFiles2);
                //Before filter
                powerpointImage.Add(GenImageByIndividual(profileItem));

                //After filter
                if (_profileToolForm.checkBoxOnlyLast.Checked)
                {
                    bool flag = false;
                    for (int i = 0; i < _loopCount; i++)
                    {
                        bool isFilter = profileReader.Filter(profileItem, pulseWidth, stdevSpec);
                        if (isFilter)
                            flag = true;
                    }
                    if (profileItem.Value.Any() && _loopCount != 0)
                        powerpointImage.Add(GenImageByIndividual(profileItem, _loopCount, flag));
                }
                else
                {
                    for (int i = 0; i < _loopCount; i++)
                    {
                        bool isFilter = profileReader.Filter(profileItem, pulseWidth, stdevSpec);
                        if (profileItem.Value.Any())
                            powerpointImage.Add(GenImageByIndividual(profileItem, i + 1, isFilter));
                    }
                }

                _powerpointImageDic.Add(index + 1, powerpointImage);
                if (_loopCount != 0)
                    profileReader.ExportProfileFiles(_profileFiles, _profilePath, _tempPath);

                _profileFiles.Add(profileItem);
            }
        }

        public PowerpointImage GenImageByIndividual(Profile profileFile, int loopCount = 0, bool isFilter = false)
        {
            var site = profileFile.Site;
            var pin = profileFile.Pin;
            var name = pin + "_Site" + site;
            using (Chart chart1 = new Chart())
            {
                chart1.Name = name;
                chart1.Height = PowerpointHeight;
                chart1.Width = PowerpointWidth;
                Title title = chart1.Titles.Add(name);
                title.Font = new Font("Arial", 28, FontStyle.Regular);
                if (isFilter)
                    title.ForeColor = Color.Red;
                ChartArea area = chart1.ChartAreas.Add("chartArea");
                area.BorderDashStyle = ChartDashStyle.Solid;
                area.AxisX.Title = "Time(s)";
                area.AxisX.TitleFont = new Font("Arial", 28, FontStyle.Regular);
                area.AxisY.MajorTickMark.TickMarkStyle = TickMarkStyle.None;
                area.AxisY.Title = profileFile.ChartType;
                area.AxisY.TitleFont = new Font("Arial", 28, FontStyle.Regular);
                if (_profileToolForm.checkBox_Power.Checked && profileFile.ChartType != "Power")
                    area.AxisY.TitleForeColor = Color.Red;

                area.Position.X = 10;
                area.Position.Y = 15;
                area.Position.Height = 84;
                area.Position.Width = 80;
                area.BackColor = Color.Black;

                Legend legend = new Legend();
                legend.Docking = Docking.Top;
                legend.Alignment = StringAlignment.Center;
                legend.Position.Auto = true;
                legend.Font = new Font("Arial", 12, FontStyle.Regular);
                legend.BackColor = Color.Transparent;

                chart1.Legends.Add(legend);

                int dateCnt = 1;
                List<double> xAxis = new List<double>();
                List<double> yAxis = profileFile.Value;
                foreach (var value in profileFile.Value)
                {
                    xAxis.Add(dateCnt / profileFile.SampleRate);
                    dateCnt++;
                }
                Series series = chart1.Series.Add(profileFile.Item);
                series.ChartType = SeriesChartType.Line;
                series.Color = Color.Aqua;
                string fileName;
                if (loopCount == 0)
                    fileName = Path.GetFileNameWithoutExtension(profileFile.FilePath);
                else
                    fileName = Path.GetFileNameWithoutExtension(profileFile.FilePath) + "_Fiter_" + loopCount;

                series.LegendText = fileName;
                series.Points.DataBindXY(xAxis, yAxis);
                area.RecalculateAxesScale();

                var imageFile = _tempPath + "\\" + fileName + ".png";// +"_" + DateTime.Now.ToString("ddHHmmssffffff") + ".png";
                _profileToolForm.AppendText(string.Format("Exporting file {0} ...", Path.GetFileNameWithoutExtension(imageFile)), Color.Blue);
                chart1.SaveImage(imageFile, ChartImageFormat.Png);
                PowerpointImage powerpointImage = new PowerpointImage();
                powerpointImage.ImageFile = imageFile;
                powerpointImage.ProfileFile = profileFile;
                return powerpointImage;
            }
        }

        public Dictionary<int, List<PowerpointImage>> GenImageMerge(List<Profile> profileFiles)
        {
            Dictionary<int, List<PowerpointImage>> dic = new Dictionary<int, List<PowerpointImage>>();
            var pinGroups = profileFiles.GroupBy(x => x.Pin).ToList();
            int groupCnt = 0;
            foreach (var pinGroup in pinGroups)
            {
                var pin = pinGroup.First().Pin;
                var siteGroups = pinGroup.GroupBy(x => x.Site);
                List<PowerpointImage> images = new List<PowerpointImage>();

                foreach (var siteGroup in siteGroups)
                {
                    var sortSiteGroup = siteGroup.ToList().OrderBy(x => x.Date);
                    var currentChartCount = sortSiteGroup.Sum(x => x.Value.Count);
                    var site = sortSiteGroup.First().Site;
                    var name = pin + "_Site" + site;

                    using (Chart chart = new Chart())
                    {
                        if (_profileToolForm.checkBox_Legend.Checked)
                        {
                            Legend legend = new Legend();
                            legend.Docking = Docking.Top;
                            legend.Alignment = StringAlignment.Center;
                            legend.AutoFitMinFontSize = 8;
                            legend.Position.Auto = true;
                            legend.Font = new Font("Arial", 12, FontStyle.Regular);
                            legend.BackColor = Color.Transparent;
                            chart.Legends.Add(legend);
                        }
                        #region add chart Series
                        double startTime = 0;
                        if (currentChartCount < MaxChartCount)
                        {
                            for (int i = 0; i < sortSiteGroup.Count(); i++)
                            {
                                var row = sortSiteGroup.ElementAt(i);
                                int dateCnt = 1;
                                Series series = new Series(row.Item + i);
                                series.ChartType = SeriesChartType.Line;
                                series.LegendText = Path.GetFileNameWithoutExtension(row.FilePath);
                                chart.Series.Add(series);
                                List<double> xAxis = new List<double>();
                                List<double> yAxis = row.Value;
                                foreach (var value in row.Value)
                                {
                                    xAxis.Add(startTime + dateCnt / row.SampleRate);
                                    dateCnt++;
                                }
                                series.Points.DataBindXY(xAxis, yAxis);
                                startTime += dateCnt / row.SampleRate;
                            }
                        }
                        else
                        {
                            _chartMergeCount = currentChartCount / DevideChartCount;
                            _chartMergeCount = _chartMergeCount == 0 ? 1 : _chartMergeCount;
                            for (int i = 0; i < sortSiteGroup.Count(); i++)
                            {
                                var row = sortSiteGroup.ElementAt(i);
                                chart.Series.Add(MergeSeries(row, row.Item + i, ref startTime, false));
                            }
                        }

                        #endregion
                        SetChartFormat(chart, name);
                        var imageFile = _tempPath + "\\" + name + "_" + DateTime.Now.ToString("yyyyMMddHHmmssffffff") + ".png";
                        _profileToolForm.AppendText(string.Format("Exporting file {0} ...", Path.GetFileNameWithoutExtension(imageFile)), Color.Blue);
                        chart.SaveImage(imageFile, ImageFormat.Png);
                        chart.Dispose();
                        PowerpointImage powerpointImage = new PowerpointImage();
                        powerpointImage.ImageFile = imageFile;
                        images.Add(powerpointImage);
                    }
                }
                groupCnt++;
                dic.Add(groupCnt, images);
            }
            return dic;
        }

        private Series MergeSeries(Profile row, string name, ref double startTime, bool isChangestartTime = true)
        {
            Series series = new Series(name);
            series.ChartType = SeriesChartType.Line;
            series.LegendText = Path.GetFileNameWithoutExtension(row.FilePath);
            int dateCnt = 0;
            int startIndex = 0;
            List<double> xAxis = new List<double>();
            List<double> yAxis = new List<double>();
            while (startIndex < row.Value.Count)
            {
                xAxis.Add(startTime + dateCnt * (_chartMergeCount / 2) / row.SampleRate);
                yAxis.Add(row.Value[startIndex]);
                dateCnt++;
                startIndex = startIndex + _chartMergeCount;
            }
            series.Points.DataBindXY(xAxis, yAxis);
            startTime += dateCnt / row.SampleRate;

            return series;
        }

        private void SetChartFormat(Chart chart, string name)
        {
            chart.Name = name;
            chart.Height = _profileToolForm.checkBox_Legend.Checked ? MultiPinHeight : PowerpointHeight;
            chart.Width = PowerpointWidth;
            Title title = chart.Titles.Add(name);
            title.Font = new Font("Arial", 28, FontStyle.Regular);

            ChartArea area = chart.ChartAreas.Add("chartArea");
            area.BorderDashStyle = ChartDashStyle.Solid;
            area.AxisX.Title = "Time(s)";
            area.AxisX.TitleFont = new Font("Arial", 28, FontStyle.Regular);
            area.AxisY.MajorTickMark.TickMarkStyle = TickMarkStyle.None;
            area.AxisY.TitleFont = new Font("Arial", 28, FontStyle.Regular);

            area.AxisX.MajorGrid.Enabled = false;
            area.AxisY.IntervalAutoMode = IntervalAutoMode.VariableCount;
            area.Position.X = 10;
            area.Position.Y = 15;
            area.Position.Height = 84;
            area.Position.Width = 80;
            area.BackColor = Color.White;

            if (_profileToolForm.checkBox_Legend.Checked)
            {
                Legend legend = new Legend();
                legend.Docking = Docking.Top;
                legend.Alignment = StringAlignment.Center;
                legend.AutoFitMinFontSize = 8;
                legend.Position.Auto = true;
                legend.Font = new Font("Arial", 12, FontStyle.Regular);
                legend.BackColor = Color.Transparent;
                chart.Legends.Add(legend);
            }

            area.AxisX.RoundAxisValues();
            area.AxisX.IsStartedFromZero = true;
            area.AxisX.Minimum = 0;
            var values = chart.Series.SelectMany(x => x.Points).OrderBy(x => x.XValue).Last().XValue;
            area.AxisX.Maximum = Math.Round(values, 3);
            area.RecalculateAxesScale();
        }
    }
}