using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Reflection;
using System.IO;
using DevExpress.XtraRichEdit;
using System.Xml.Linq;
using System.Dynamic;

namespace ExpandoMailMerge
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
            this.Loaded += new RoutedEventHandler(MainPage_Loaded);

        }

        void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            richEditControl1.ApplyTemplate();

            dynamic weathers = GetExpandoFromXml("weather.xml");
            richEditControl1.Options.MailMerge.DataSource = weathers;
            
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExpandoMailMerge.weather_report.rtf");
            richEditControl1.LoadDocument(stream, DocumentFormat.Rtf);
            
            richEditControl1.Options.MailMerge.ViewMergedData = true;
        }

        public static IList<dynamic> GetExpandoFromXml(String file)
        {
            var weathers = new List<dynamic>();

            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExpandoMailMerge" + "." + file);
            var doc = XDocument.Load(stream);
            var nodes = from node in doc.Root.Descendants("weather")
                        select node;
            foreach (var n in nodes) {
                dynamic MyData = new ExpandoObject();
                MyData.LastUpdateTime = String.Format("{0:o}", DateTime.Now);
                MyData.Weather = new ExpandoObject();
                foreach (var child in n.Descendants()) {

                    var w = MyData.Weather as IDictionary<String, object>;
                    XAttribute atb = child.Attribute("data");
                    if (atb != null)
                        w[child.Name.LocalName] = atb.Value;

                }

                weathers.Add(MyData);

            }
            return weathers;
        }

    }
}
