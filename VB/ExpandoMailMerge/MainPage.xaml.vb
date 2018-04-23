Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Net
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes
Imports System.Reflection
Imports System.IO
Imports DevExpress.XtraRichEdit
Imports System.Xml.Linq
Imports System.Dynamic

Namespace ExpandoMailMerge
	Partial Public Class MainPage
		Inherits UserControl
		Public Sub New()
			InitializeComponent()
			AddHandler Loaded, AddressOf MainPage_Loaded

		End Sub

		Private Sub MainPage_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			richEditControl1.ApplyTemplate()
            Dim weathers As Object = GetExpandoFromXml("weather.xml")
			richEditControl1.Options.MailMerge.DataSource = weathers

			Dim stream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("weather_report.rtf")
            richEditControl1.LoadDocument(stream, DocumentFormat.Rtf)

			richEditControl1.Options.MailMerge.ViewMergedData = True
		End Sub

        Public Shared Function GetExpandoFromXml(ByVal file As String) As IList(Of Object)
            Dim weathers = New List(Of Object)()

            Dim stream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("ExpandoMailMerge" & "." & file)
            Dim doc = XDocument.Load(stream)
            Dim nodes = _
             From node In doc.Root.Descendants("weather") _
             Select node
            For Each n In nodes
                'INSTANT VB TODO TASK: There is no VB equivalent to the C# 'dynamic' keyword:
                'ORIGINAL LINE: dynamic MyData = New ExpandoObject();
                Dim MyData As Object = New ExpandoObject()
                MyData.LastUpdateTime = String.Format("{0:o}", DateTime.Now)
                MyData.Weather = New ExpandoObject()
                For Each child In n.Descendants()

                    Dim w = TryCast(MyData.Weather, IDictionary(Of String, Object))
                    Dim atb As XAttribute = child.Attribute("data")
                    If atb IsNot Nothing Then
                        w(child.Name.LocalName) = atb.Value
                    End If

                Next child

                weathers.Add(MyData)

            Next n
            Return weathers
        End Function

	End Class
End Namespace
