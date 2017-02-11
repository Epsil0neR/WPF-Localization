Imports System.Globalization

Public Class Application
    Private Shared m_Languages As New List(Of CultureInfo)

    Public Shared ReadOnly Property Languages As List(Of CultureInfo)
        Get
            Return m_Languages
        End Get
    End Property

    Public Sub New()
        InitializeComponent()

        m_Languages.Clear()
        m_Languages.Add(New CultureInfo("en-US")) 'Нейтральная культура для этого проекта
        m_Languages.Add(New CultureInfo("ru-RU"))

        Language = My.Settings.DefaultLanguage
    End Sub

    'Евент для оповещения всех окон приложения
    Public Shared Event LanguageChanged(sender As Object, e As EventArgs)

    Public Shared Property Language As CultureInfo
        Get
            Return System.Threading.Thread.CurrentThread.CurrentUICulture
        End Get
        Set(value As CultureInfo)
            If value Is Nothing Then Throw New ArgumentNullException("value")
            If value.Equals(System.Threading.Thread.CurrentThread.CurrentUICulture) Then Exit Property

            '1. Меняем язык приложения:
            System.Threading.Thread.CurrentThread.CurrentUICulture = value

            '2. Создаём ResourceDictionary для новой культуры
            Dim dict As New ResourceDictionary()
            Select Case value.Name
                Case "ru-RU"
                    dict.Source = New Uri(String.Format("Resources/lang.{0}.xaml", value.Name), UriKind.Relative)
                Case Else
                    dict.Source = New Uri("Resources/lang.xaml", UriKind.Relative)
            End Select

            '3. Находим старую ResourceDictionary и удаляем его и добавляем новую ResourceDictionary
            Dim oldDict As ResourceDictionary = (From d In My.Application.Resources.MergedDictionaries
                                                 Where d.Source IsNot Nothing _
                                                 AndAlso d.Source.OriginalString.StartsWith("Resources/lang.")
                                                 Select d).FirstOrDefault
            If oldDict IsNot Nothing Then
                Dim ind As Integer = My.Application.Resources.MergedDictionaries.IndexOf(oldDict)
                My.Application.Resources.MergedDictionaries.Remove(oldDict)
                My.Application.Resources.MergedDictionaries.Insert(ind, dict)
            Else
                My.Application.Resources.MergedDictionaries.Add(dict)
            End If

            '4. Вызываем евент для оповещения всех окон.
            RaiseEvent LanguageChanged(Application.Current, New EventArgs)
        End Set
    End Property

    Private Shared Sub OnLanguageChanged(sender As Object, e As EventArgs) Handles MyClass.LanguageChanged
        My.Settings.DefaultLanguage = Language
        My.Settings.Save()
    End Sub

End Class
