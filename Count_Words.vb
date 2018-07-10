Imports System.IO
Module Count_Words
    'Skapad av Markus Nordin
    'Kod Utmaning Count Words
    'Läs in från fil och skapa en sammanfattning
    'd:\test.txt (3764 ord enligt Google Docs)
    '3835 enligt programmet, men då räknas siffror
    'IsNumeric ska användas för att undvika siffror
    Sub Main()
        Dim Sökväg As String = "d:\test.txt"
        Dim testString As String = LäsInFil(Sökväg)
        Dim Buffer As String = ""
        Dim MaxAntalOrd As Integer = 0
        Dim AntalOrd As Integer = 0
        Dim Bokstav As Char = ""
        Dim Ord As New List(Of String)
        For Each Bokstav In testString
            Select Case Bokstav 'Filter för strängen. Tar bort det vi inte vill ha
                Case " "
                    If Buffer = "" Then
                    Else
                        If Char.IsNumber(Buffer) Then
                            Buffer = ""
                        Else
                            Ord.Add(Buffer)
                            Buffer = ""
                        End If
                    End If
                Case "-"
                Case "#"
                Case Else
                    Buffer = Buffer & Bokstav
            End Select
        Next
        'Sista ordet i dokumentet
        If Buffer = "" Then
        Else
            Ord.Add(Buffer)
            Buffer = ""
        End If
        AntalOrd = Ord.Count 'Räknar antal ord
        Console.WriteLine("Antal ord är " & AntalOrd)
        Call SkrivFil("d:\test2.txt", Ord) 'skriver ut en testfil med informationen i Array Ord
        Console.ReadKey()
    End Sub
    Function LäsInFil(ByVal Filnamn As String)
        Dim Fil As String = ""
        Dim Buffer As String = ""
        Dim Sökväg As String = Filnamn

        If File.Exists(Sökväg) = False Then
            Console.WriteLine("Hittade inte filen")
        Else
            Try
                Buffer = File.ReadAllText(Sökväg)
            Catch ex As Exception
                'Console.WriteLine(ex)
                Console.WriteLine("Något gick fel")
            End Try
        End If
        Fil = Buffer.Replace(vbCrLf, " ") 'Tar bort ENTER slag och ersätter med ett mellanslag.
        Buffer = Fil
        Fil = Buffer.Replace(vbTab, " ")
        Return Fil
    End Function
    Sub SkrivFil(ByVal Filnamn As String, ByVal testString As IList(Of String))
        Dim Sökväg As String = Filnamn
        Dim Buffer As List(Of String) = testString
        File.WriteAllLines(Sökväg, Buffer)
    End Sub
End Module
