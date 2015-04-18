Imports System.Globalization

Class MainWindow

    Public Sub New()
        InitializeComponent()

        'Добавляем обработчик события смены языка у приложения
        AddHandler Application.LanguageChanged, AddressOf LanguageChanged

        Dim currLang = Application.Language

        'Заполняем меню смены языка:
        menuLanguage.Items.Clear()
        For Each lang In Application.Languages
            Dim menuLang As New MenuItem()
            menuLang.Header = lang.DisplayName
            menuLang.Tag = lang
            menuLang.IsChecked = lang.Equals(currLang)
            AddHandler menuLang.Click, AddressOf ChangeLanguageClick
            menuLanguage.Items.Add(menuLang)
        Next
    End Sub

    Private Sub LanguageChanged(sender As Object, e As EventArgs)
        Dim currLang = Application.Language

        'Отмечаем нужный пункт смены языка как выбранный язык
        For Each i As MenuItem In menuLanguage.Items
            Dim ci As CultureInfo = TryCast(i.Tag, CultureInfo)
            i.IsChecked = ci IsNot Nothing AndAlso ci.Equals(currLang)
        Next
    End Sub

    Private Sub ChangeLanguageClick(sender As Object, e As RoutedEventArgs)
        Dim mi As MenuItem = TryCast(sender, MenuItem)
        If mi IsNot Nothing Then
            Dim lang As CultureInfo = TryCast(mi.Tag, CultureInfo)
            If lang IsNot Nothing Then
                Application.Language = lang
            End If
        End If
    End Sub

End Class
