using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;
using System.Windows.Forms;
using craigslist.org.Models;
using craigslist.org.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using Color = System.Drawing.Color;

namespace craigslist.org
{
    public partial class MainForm : MetroForm
    {
        public bool LogToUi = true;
        public bool LogToFile = true;
        private readonly string _path = Application.StartupPath;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        public Dictionary<string, int> _sheetsIds = new Dictionary<string, int>();
        Random _rnd = new Random();
        //public static SheetsService Service;
        private readonly List<object> _headers = new List<object>
        {
            "Tile",
            "Price",
            "Location",
            "Date and time",
            "Description",
            "Url"
        };

        public MainForm()
        {
            InitializeComponent();
        }


        private async Task MainWork()
        {



            var categories = JsonConvert.DeserializeObject<List<Category>>(File.ReadAllText("js1.txt"));
            //var isSheetExist = await GoogleSheetApiService.CheckIfSheetExist(GoogleSheetIdI.Text, categories.First().CategoryName.Replace(":", ""));
            var urlCategories = File.ReadAllLines(inputI.Text);
            //var categories = new List<Category>();
            foreach (var urlCategory in urlCategories)
            {
                //if (urlCategory.Trim() == "")
                //{
                //    continue;
                //}
                //var doc = await HttpCaller.GetDoc(urlCategory);
                //var categoryName = doc.DocumentNode.SelectSingleNode("//select[@id='areaAbb']/option[1]").InnerText +
                //                   "/" + doc.DocumentNode.SelectSingleNode("//select[@id='catAbb']/option[@selected]").InnerText +
                //                   "/" + doc.DocumentNode.SelectSingleNode("//select[@id='subcatAbb']/option[@selected]")
                //                       .InnerText;
                //var realEstates = doc.DocumentNode.SelectNodes("//li[@data-pid]/div");
                //if (realEstates == null)
                //{
                //    ErrorLog($"the following real estate category is empty ==> {urlCategory}");
                //    continue;
                //}

                //var urls = await GetPages(urlCategory);

                //var threads = int.Parse(threadsI.Text);
                //var tpl = new TransformBlock<string, List<RealEstate>>(async x => await ScrapeRealEstateInfo(x).ConfigureAwait(false),
                //    new ExecutionDataflowBlockOptions { MaxDegreeOfParallelism = threads });

                //foreach (var url in urls)
                //{
                //    tpl.Post(url);
                //}

                //var category = new Category { CategoryName = categoryName };
                //foreach (var url in urls)
                //{
                //    var listOfRealEstate = await tpl.ReceiveAsync().ConfigureAwait(false);
                //    if (listOfRealEstate != null)
                //    {
                //        category.RealEstate.AddRange(listOfRealEstate);
                //    }
                //}
                //categories.Add(category);

            }
            //var json = JsonConvert.SerializeObject(categories, Formatting.Indented);
            //File.WriteAllText("js1.txt", json);
            //categories = JsonConvert.DeserializeObject<List<Category>>(File.ReadAllText("json1.txt"));

            foreach (var ctg in categories)
            {
                var values = new List<IList<object>>();

                var categoryName = ctg.CategoryName.Replace(":", "");


                try
                {
                    await GoogleSheetApiService.DeleteRows(GoogleSheetIdI.Text, categoryName,_sheetsIds[categoryName]);
                }
                catch (Exception)
                {

                    var sheetId = await GoogleSheetApiService.CreateNewSheet(GoogleSheetIdI.Text, categoryName);
                    _sheetsIds[categoryName] = sheetId;
                    values.Add(_headers);
                }

              

                //foreach (var realEstate in ctg.RealEstate)
                //{
                //    var obj = new List<object>
                //    {
                //        realEstate.Title,
                //        realEstate.Price,
                //        realEstate.Location,
                //        realEstate.DateTime,
                //        realEstate.Description,
                //        realEstate.Url
                //    };
                //    values.Add(obj);
                //    if (values.Count == 3000)
                //    {
                //        await GoogleSheetApiService.AppendData(values, GoogleSheetIdI.Text, categoryNmae);
                //        values = new List<IList<object>>();
                //    }
                //}
                await GoogleSheetApiService.AppendData(values, GoogleSheetIdI.Text, categoryName);

            }
        }

        private async Task<List<string>> GetPages(string urlCategory)
        {
            var urls = new List<string>();

            var page = 0;
            do
            {
                var urlPage = urlCategory + "?s=" + page;
                var doc = await HttpCaller.GetDoc(urlPage);
                var nextPage = doc.DocumentNode.SelectSingleNode("//a[@title='next page']")
                    ?.GetAttributeValue("href", "") ?? "";
                if (string.IsNullOrEmpty(nextPage))
                {
                    urls.Add(urlPage);
                    break;
                }
                page = page + 120;
                urls.Add(urlPage);
            } while (true);

            return urls;
        }

        private async Task<List<RealEstate>> ScrapeRealEstateInfo(string url)
        {
            var realEstates = new List<RealEstate>();
            var doc = new HtmlAgilityPack.HtmlDocument();
            var urlRealEstate = "";
            try
            {
                urlRealEstate = url;
                doc = await HttpCaller.GetDoc(url);
                var realEstatesNodes = doc.DocumentNode.SelectNodes("//li[@data-pid]/div");
                foreach (var realEstate in realEstatesNodes)
                {
                    var dateTime = realEstate.SelectSingleNode(".//time").GetAttributeValue("datetime", "");
                    var title = realEstate.SelectSingleNode(".//a")?.InnerText ?? "N/A";
                    urlRealEstate = realEstate.SelectSingleNode("./h3//a").GetAttributeValue("href", "");
                    var price = realEstate.SelectSingleNode(".//span[@class='result-price']")?.InnerText.Trim() ?? "N/A";
                    var location =
                        realEstate.SelectSingleNode(".//span[@class='result-hood']")?.InnerText.Trim().Replace("(", "")
                            .Replace(")", "") ?? realEstate.SelectSingleNode(".//span[@class='nearby']")?.InnerText
                            .Trim().Replace("(", "").Replace(")", "") ?? "N/A";
                    doc = await HttpCaller.GetDoc(urlRealEstate);
                    var description = doc.DocumentNode.SelectSingleNode("//section[@id='postingbody']")?.InnerText.Trim() ?? "N/A";
                    if (price == "N/A")
                    {
                        price = await GetPrice(title);
                    }
                    realEstates.Add(new RealEstate { Location = location, Description = description, Title = title, DateTime = dateTime, Url = urlRealEstate, Price = price });
                }
            }
            catch (Exception)
            {
                var count = _rnd.Next(0, 999999999);
                var newHtml = "<!--" + urlRealEstate + "-->" + doc.DocumentNode.OuterHtml;
                doc.LoadHtml(newHtml);
                doc.Save($@"Failed scraping/New html {count}.html");
                return null;
            }

            return realEstates;
        }

        private async Task<string> GetPrice(string title)
        {
            var price = "";
            if (title.Contains("$"))
            {
                var x = title.IndexOf("$", StringComparison.Ordinal);

                if (x != title.Length - 1 && title[x + 1] != ' ')
                {
                    var chars = new StringBuilder();
                    do
                    {

                        var check = char.IsDigit(title[x + 1]);
                        if (check)
                        {
                            chars.Append(title[x + 1]);
                        }
                        else
                        {
                            break;
                        }
                        x++;
                        if (x + 1 == title.Length)
                        {
                            break;
                        }
                    } while (true);
                    price = "$" + chars.ToString();
                }
                x = title.IndexOf("$", StringComparison.Ordinal);
                var index = x;
                if (x > 0 && char.IsDigit(title[x - 1]))
                {
                    var chars = new StringBuilder();
                    do
                    {

                        var check = char.IsDigit(title[x - 1]);
                        if (check)
                        {
                            chars.Append(title[x - 1]);
                        }
                        else
                        {
                            break;
                        }
                        x--;
                        if (x == 0)
                        {
                            break;
                        }
                    } while (true);
                    price = "$" + chars.ToString();
                    if (index == title.Length - 1)
                    {
                        price = new string(Enumerable.Range(1, price.Length)
                            .Select(i => price[price.Length - i]).ToArray());
                    }
                }
            }

            return price;
        }

        private async Task<RealEstate> EntittyTest(string url)
        {
            var doc = await HttpCaller.GetDoc(url);

            var description = doc.DocumentNode.SelectSingleNode("//section[@id='postingbody']")?.InnerText.Trim() ?? "N/A";
            var title = doc.DocumentNode.SelectSingleNode("//span[@id='titletextonly']")?.InnerText.Trim() ?? "N/A";
            var price = "N/A";
            if (price == "N/A")
            {
                if (title.Contains("$"))
                {
                    var x = title.IndexOf("$", StringComparison.Ordinal);
                    if (title[x + 1] != ' ')
                    {
                        var chars = new StringBuilder();
                        do
                        {

                            var check = char.IsDigit(title[x + 1]);
                            if (check)
                            {
                                chars.Append(title[x + 1]);
                            }
                            else
                            {
                                //if (title[x + 1] == '/' && char.IsDigit(title[x + 2])|| title[x + 1] == ',' && char.IsDigit(title[x + 2]))
                                //{
                                //    chars.Append(title[x + 1]);
                                //}
                                //else if (title[x + 1] == '-' && title[x+2]=='$'&& char.IsDigit(title[x + 3]))
                                //{
                                //    chars.Append("-$");
                                //    x += 2;
                                //    continue;
                                //}
                                //else
                                //{
                                break;
                                //}
                            }
                            x++;
                            if (x + 1 == title.Length)
                            {
                                break;
                            }
                        } while (true);
                        price = "$" + chars.ToString();
                    }
                    x = title.IndexOf("$", StringComparison.Ordinal);
                    if (char.IsDigit(title[x - 1]))
                    {
                        var chars = new StringBuilder();
                        do
                        {

                            var check = char.IsDigit(title[x - 1]);
                            if (check)
                            {
                                chars.Append(title[x - 1]);
                            }
                            else
                            {
                                break;
                            }
                            x--;
                            if (x == 0)
                            {
                                break;
                            }
                        } while (true);
                        price = "$" + chars.ToString();
                    }
                }
            }


            return null;
        }


        private async void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists("Failed scraping"))
            {
                Directory.CreateDirectory("Failed scraping");
            }
            if (File.Exists("Sheets Ids.txt"))
            {
                _sheetsIds = JsonConvert.DeserializeObject<Dictionary<string, int>>(File.ReadAllText("Sheets Ids.txt"));
            }
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            LoadConfig();
            //await GoogleSheetApiService.Credential();

        }

        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_config[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_config[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _config[x.Name];
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_config[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _config.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _config.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _config = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_config, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _config = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (LogToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (LogToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
        }

        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
        }
        private void loadInputB_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog { Filter = @"TXT|*.txt", InitialDirectory = _path };
            if (o.ShowDialog() == DialogResult.OK)
            {
                inputI.Text = o.FileName;
            }
        }
        private void openInputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                //Process.Start(inputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                //Process.Start(outputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            var saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"csv file|*.csv",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                //outputI.Text = saveFileDialog1.FileName;
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            //var rr = await ScrapeRealEstateInfo("https://cnj.craigslist.org/d/wanted%3A-room-share/search/sha?s=0");
            //var t = await EntittyTest("https://newyork.craigslist.org/que/sha/d/1100-lovely-room-rent-all-furnished/7318573080.html");
            //var t = await EntittyTest("https://cnj.craigslist.org/sha/d/mature-working-class-male-in-need-of/7319181427.html");
            //return;
            Display("Start scraping....");
            //await GoogleSheetApiService.Credential();
            var p = 0;
            do
            {
                await GoogleSheetApiService.Credential();
                await Task.Run(MainWork);
                p += 1;
                Display("Entries has been updated " + p + " times");
                await Task.Delay(10000);
            } while (true);
        }



        //private void checkBox1_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (NewGoogleSheet.Checked)
        //    {
        //        GoogleSheetName.Visible = true;
        //        label2.Visible = true;
        //    }
        //    else
        //    {
        //        GoogleSheetName.Visible = false;
        //        label2.Visible = false;
        //    }
        //}


    }
}
