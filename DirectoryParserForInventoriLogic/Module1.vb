Imports Inventor
Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
'Imports System.Xml.Linq

Module Module1
#Region "Working Console App Code"
    ''' <summary>
    ''' The Main Program.
    ''' Could modify this app to use .NET 6 features such as the XML to LINQ mechanic but I have no idea if Windows 7 even supports it...?
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Main()
        Console.WriteLine("Enter Path to Parse...")
        Dim rootPath As String = Console.ReadLine()
        Dim dir = New DirectoryInfo(rootPath)
        Dim doc = New XDocument(GetDirectoryXml(dir, 0))
        doc.Save("C:\temp\VBtest.xml")
        Console.WriteLine("Done")
        'Console.WriteLine(doc.ToString())

        'Console.Read()
    End Sub
    ''' <summary>
    ''' Gets a friendly-named ParentAssembly, Assembly and level for files in dir
    ''' </summary>
    ''' <param name="dir">the directory to parse</param>
    ''' <param name="level">the current level</param>
    ''' <returns>XElements for the resultant XML file</returns>
    ''' <remarks>Need to edit the resultant .XML file with notepad++ using the Regex search pattern of "<dir dirname=".*">" and "</dir>" to remove excess information</remarks>
    Public Function GetDirectoryXml(ByVal dir As DirectoryInfo, ByVal level As Long) As Object
        Dim info = New XElement("dir", New XAttribute("dirname", GetFriendlyDirName(dir.Name)))
        If Not dir.Name.Contains("Superseded") Then
            For Each file As FileInfo In dir.GetFiles()
                'info.Add(New XElement("file", New XAttribute("name", file.Name), New XAttribute("friendlyname", GetFriendlyName(file.Name))))
                If Not file.Name.Contains("IL") And Not file.Name.Contains("DL") And Not file.Name.Contains("SP") Then
                    'if the directory name is the same as the assembly name then the parentassembly is the folder above!
                    If GetFriendlyDirName(dir.Name) = GetFriendlyName(file.Name) Then
                        If getsheetnum(file.Name) <= 1 Then
                            info.Add(New XElement("file",
                                                  New XAttribute("parentassembly", GetFriendlyDirName(dir.Parent.Name)),
                                                  New XAttribute("friendlyname", GetFriendlyName(file.Name)),
                                                  New XAttribute("Level", level + 1)))
                        End If
                    Else
                        If getsheetnum(file.Name) <= 1 Then
                            info.Add(New XElement("file",
                                                  New XAttribute("parentassembly", GetFriendlyDirName(dir.Name)),
                                                  New XAttribute("friendlyname", GetFriendlyName(file.Name)),
                                                  New XAttribute("level", level + 2)))
                        End If
                    End If
                End If
            Next

            For Each subDir As DirectoryInfo In dir.GetDirectories()
                If Not subDir.Name.Contains("Superseded") Then
                    info.Add(GetDirectoryXml(subDir, level + 1))
                End If
            Next
        End If
        Return info

    End Function

    ''' <summary>
    ''' Uses Regex to get a nicely formatted filename
    ''' </summary>
    ''' <param name="p">The string to search</param>
    ''' <returns>returns a string formatted thus: @@-#####-000 or @@-@#####-000</returns>
    ''' <remarks></remarks>
    Public Function GetFriendlyName(p As String) As Object
        Dim f As String = String.Empty
        Dim r As New Regex("\w{2}-\d{5,}|\w{2}-\w\d{5,}")
        f = r.Match(p).Captures(0).ToString() + "-000"
        Console.WriteLine(f)
        Return f
    End Function

    ''' <summary>
    ''' Uses Regex to get a nicely formatted dirname
    ''' </summary>
    ''' <param name="p1">The string to search</param>
    ''' <returns>returns a string formatted thus: @@-#####-000 or @@-@#####-000</returns>
    ''' <remarks></remarks>
    Public Function GetFriendlyDirName(p1 As String) As Object
        If Not p1.Contains(":") Then
            Dim f As String = String.Empty
            Dim r As New Regex("\d{3,}|\w\d{3,}")
            f = "AS-" + r.Match(p1).Captures(0).ToString() + "-000"
            Return f
        Else
            Return p1
        End If
    End Function

    ''' <summary>
    ''' Searches a given string to see if it matches the required pattern.
    ''' </summary>
    ''' <param name="p1">the string to query</param>
    ''' <returns>Returns an int value</returns>
    ''' <remarks></remarks>
    Private Function getsheetnum(p1 As String) As Integer
        Dim f As String = String.Empty
        Dim pattern As String = "(.*)(sht-)(\d{3})(.*)"
        Dim matches As MatchCollection = Regex.Matches(p1, pattern)
        For Each m As Match In matches
            Dim g As Group = m.Groups(3)
            f = CInt(g.Value)
        Next
        Return CInt(f)
    End Function
#End Region


End Module