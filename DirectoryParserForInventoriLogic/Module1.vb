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
    Public Sub Main()
        Console.WriteLine("Enter Path to Parse...")

        Dim rootPath As String = Console.ReadLine()
        Dim dir = New DirectoryInfo(rootPath)
        Dim doc = New XDocument(GetDirectoryXml(dir, 0))
        doc.Save("C:\temp\VBtest.xml")
        Console.WriteLine(doc.ToString())

        Console.Read()
    End Sub

    Public Function GetDirectoryXml(ByVal dir As DirectoryInfo, ByVal level As Long) As Object
        Dim info = New XElement("dir", New XAttribute("name", dir.Name))
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

    Public Function GetFriendlyName(p As String) As Object
        Dim f As String = String.Empty
        Dim r As New Regex("\w{2}-\d{5,}|\w{2}-\w\d{5,}")
        f = r.Match(p).Captures(0).ToString() + "-000"
        Console.WriteLine(f)
        Return f
    End Function

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


    'Dim ThisApplication As Application
    'Public PartsList As List(Of SubObjectCls)
    'Public r As List(Of SubObjectCls)
    'Public ParentList As List(Of String)

    'Public Sub BeginCreateAssemblyStructure()
    '    'define the parent assembly
    '    Dim asmDoc As AssemblyDocument
    '    asmDoc = ThisApplication.ActiveDocument
    '    Dim filename As String = System.IO.Path.GetFileNameWithoutExtension(ThisApplication.ActiveDocument.Displayname)
    '    Dim filetab As String = InputBox("Which project?", "4 Letter Project Code", "HATC") + "-MODELLING-BASELINE"

    '    PartsList = New List(Of SubObjectCls)
    '    Dim FilesArray As New ArrayList
    '    Dim tr As transaction

    '    tr = ThisApplication.TransactionManager.StartTransaction( _
    '        ThisApplication.ActiveDocument, _
    '        "Create Standard Parts From Excel")
    '    'this is to simply set the excel values to the correct file/tab- nothing more!
    '    FileArray = GoExcel.CellValues("C:\LEGACY VAULT WORKING FOLDER\Designs\Project Tracker.xlsx", filetab, "A3", "A4") ' sets excel to the correct sheet!

    '    'Data collection:
    '    For MyRow = 3 To 50 ' max limit = 50 rows for debugging purposes
    '        Dim SO As SubObjectCls
    '        SO = New SubObjectCls
    '        'not sure if we should change this to Column C as it contains the files we know about from the Vault
    '        'if we did we could then have it insert that file if we linked this routine to Vault...?
    '        If GoExcel.CellValue("A" & MyRow) = "" Then Exit For 'exits when the value is empty!
    '        Dim tmpstr As String = GoExcel.CellValue("I" & MyRow) 'parent row
    '        If Not tmpstr.StartsWith("AS-") Then
    '            Continue For
    '        End If

    '        SO.Partno = GoExcel.CellValue("B" & MyRow)  'PART NUMBER
    '        SO.LegacyDescr = GoExcel.CellValue("K" & MyRow) 'DESCRIPTION
    '        SO.LegacyRev = GoExcel.CellValue("L" & MyRow)   'REV
    '        SO.LegacyDrawingNo = GoExcel.CellValue("M" & MyRow) 'SUBJECT/LEGACY DRAWING NUMBER
    '        SO.ParentAssembly = GoExcel.CellValue("I" & MyRow)  'PARENT ASSEMBLY
    '        PartsList.Add(SO)
    '    Next
    '    MessageBox.Show("PartsList.Count= " & Partslist.Count, "Parts Count")
    '    r = New List(Of SubObjectCls)
    '    ParentList = New List(Of String)
    '    For Each a As SubObjectCls In PartsList
    '        If a.PartNo.StartsWith("AS-") And a.ParentAssembly = filename Then
    '            r.Add(a)
    '        End If
    '    Next
    '    For i = 0 To PartsList.Count - 1
    '        'copy/find components as listed in the spreadsheet:
    '        parentlist.add(PartsList(i).ParentAssembly)
    '        'MessageBox.Show(r(i).PartNo + vbCrLf + r(i).ParentAssembly,"Info in part #" + i.ToString())
    '    Next i
    '    'MessageBox.Show(parentlist.Count, "combined parent list Count")
    '    Dim FilteredParentlist = (From a As String In parentlist
    '                            Select a).Distinct()
    '    'MessageBox.Show(FilteredParentlist.Count, "combined, filtered parent list Count")
    '    'filter for assemblies with children.
    '    For i = 0 To PartsList.Count - 1
    '        If FilteredParentlist.Contains(PartsList(i).PartNo) Then
    '            PartsList(i).HasChildren = True
    '            'MessageBox.Show(PartsList(i).HasChildren,"Assembly has children")
    '            Dim children = (From a As SubObjectCls In PartsList
    '                                    Where a.ParentAssembly = PartsList(i).PartNo).ToList()
    '            PartsList(i).Children = children
    '            MessageBox.Show(PartsList(i).PartNo & " has: " & PartsList(i).Children.Count & " Children", "Assembly has children")
    '        End If
    '    Next
    '    'doesn't work!
    '    'borrowed from the "Traverse an Assembly" sample:
    '    'For Each SO As SubObjectCls In PartsList
    '    '    MessageBox.Show(SO.PartNo)
    '    '    If SO.Children Is Nothing Then ' Children.Count would naturally be zero!
    '    '        'If SO.Children.Count = 0 Then
    '    '        Call CreateAssemblyStructure(SO, SO.ParentAssembly)
    '    '    Else
    '    '        Call CreateAssemblySubStructure(SO, SO.ParentAssembly)
    '    '        'End If
    '    '    End If
    '    'Next
    '    'works but adds too many instances of sub-parts!
    '    'For i = 0 To PartsList.Count - 1
    '    '    'MessageBox.Show("Adding part: " + PartsList(i).PartNo,"Part No in part: " + i.ToString())
    '    '    If PartsList(i).PartNo = filename Then Continue For
    '    '    For j = 0 To filteredparentlist.count - 1
    '    '        If Partslist(i).ParentAssembly = filteredparentlist(j) Then
    '    '            'CreateAssemblyStructure(PartsList(i),filename)
    '    '            CreateAssemblyStructure(PartsList(i), PartsList(i).ParentAssembly)
    '    '        End If
    '    '    Next j
    '    '    'MessageBox.Show(PartsList(i).PartNo,"Part No in part: " + i.ToString())
    '    'Next

    '    For i = 0 To PartsList.Count - 1
    '        'MessageBox.Show("Adding part: " + PartsList(i).PartNo, "Part No in part: " + i.ToString())
    '        If PartsList(i).PartNo = filename Then Continue For ' as it's the current top-level assembly.
    '        Call CreateAssemblyStructure(PartsList(i), PartsList(i).ParentAssembly)
    '        '    'For j = 0 To FilteredParentlist.Count - 1
    '        '    If PartsList(i).ParentAssembly = FilteredParentlist(j) Then
    '        '        'MessageBox.Show(Partslist(i).Partno,"Part Number to add")
    '        '        'MessageBox.Show(PartsList(i).ParentAssembly,"Parent Assembly")
    '        '                If PartsList(i).Children Is Nothing Then
    '        '                    'MessageBox.Show(PartsList(i).PartNo & " has " & PartsList(i).Children.Count & " Children","Assembly has children")
    '        '                    Call CreateAssemblyStructure(PartsList(i), PartsList(i).ParentAssembly)
    '        '                Else
    '        '                    'MessageBox.Show(PartsList(i).PartNo & " has zero Children","Assembly has children")
    '        '                    Call CreateAssemblySubStructure(PartsList(i), PartsList(i).ParentAssembly)
    '        '                End If

    '        'End If
    '        '    'Next j
    '        '    'MessageBox.Show(PartsList(i).PartNo,"Part No in part: " + i.ToString())
    '    Next
    '    tr.End()
    '    InventorVb.DocumentUpdate()
    'End Sub

    'Private Sub CreateAssemblySubStructure(SO As SubObjectCls, ParentAssembly As String)
    '    MessageBox.Show("CreateAssemblySubStructureStart")
    '    For Each SubObj As SubObjectCls In SO.Children
    '        If Not SubObj.Children Is Nothing Then
    '            If SubObj.Children.Count = 0 Then
    '                Call CreateAssemblyStructure(SO, SO.ParentAssembly)
    '            Else
    '                Call CreateAssemblySubStructure(SubObj, SubObj.ParentAssembly)
    '            End If
    '        Else
    '            Call CreateAssemblyStructure(SO, SO.ParentAssembly)
    '        End If
    '    Next

    'End Sub

    'Private Function CreateAssemblyComponents(subObject As SubObjectCls) As String
    '    Dim basepartname As String = String.empty
    '    Dim newfilename As String
    '    If subObject.PartNo.StartsWith("AS-") Then
    '        newfilename = System.IO.Path.GetDirectoryName(ThisApplication.activedocument.fulldocumentname) & "\" & subObject.PartNo & ".iam"
    '        basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
    '    ElseIf subObject.PartNo.StartsWith("DT-") Then
    '        If subObject.LegacyDescr.Contains("ASSEMBLY") Or subObject.LegacyDescr.Contains("ASSY") Then
    '            newfilename = System.IO.Path.GetDirectoryName(ThisApplication.activedocument.fulldocumentname) & "\" & subObject.PartNo & ".iam"
    '            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.iam"
    '        Else
    '            newfilename = System.IO.Path.GetDirectoryName(ThisApplication.activedocument.fulldocumentname) & "\" & subObject.PartNo & ".ipt"
    '            basepartname = "C:\LEGACY VAULT WORKING FOLDER\Designs\DT-99999-000.ipt"
    '        End If
    '    End If
    '    'check if the file exists locally and copy a template to create it if not.
    '    If Not System.IO.File.Exists(newfilename) Then 'we need to create it - but we also might need to search the local working folder for it too...?
    '        MessageBox.Show("Looking for: " + newfilename, "Finding Files!")
    '        Dim tmpstr As String = FindFileInVWF(newfilename)
    '        If tmpstr = String.Empty Then
    '            'it doesn't exist anywhere else in the Local Vault Working Folder
    '            System.IO.File.Copy(basepartname, newfilename)
    '        Else
    '            newfilename = tmpstr
    '        End If
    '    End If
    '    Return newfilename
    'End Function

    'Private Sub CreateAssemblyStructure(subObject As SubObjectCls, parentName As String)
    '    'MessageBox.Show("CreateAssemblySubStructureStart")
    '    Dim asmDoc As AssemblyDocument
    '    Dim occ As ComponentOccurrence
    '    Dim occs As ComponentOccurrences
    '    Dim realOcc As ComponentOccurrence
    '    Dim realOccStr As String
    '    Dim PosnMatrix As Matrix
    '    Dim newfilename As String = CreateAssemblyComponents(subObject)

    '    PosnMatrix = ThisApplication.TransientGeometry.CreateMatrix
    '    'MessageBox.Show(subobject.PartNo, "Sub Object Part Number")
    '    If parentName = System.IO.Path.GetFileNameWithoutExtension(ThisApplication.ActiveDocument.DisplayName) Then
    '        'MessageBox.Show("parentname= " & parentname)
    '        'the parent assembly
    '        asmDoc = ThisApplication.ActiveDocument
    '        Try
    '            realOcc = asmDoc.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
    '            realOccStr = realOcc.Name
    '        Catch ex As Exception
    '            MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
    '        End Try
    '    Else
    '        'one of its children/grandchildren
    '        asmDoc = ThisApplication.ActiveDocument
    '        MessageBox.Show("Child's parentname= " & parentname)
    '        'MessageBox.Show("Sub Occurrence: " + SubObject.PartNo)
    '        Dim asmCompDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition

    '        For Each occ In asmCompDef.Occurrences
    '            'MessageBox.Show("Assembly name: " + occ.Name)
    '            If occ.Name = parentName & ":1" Then
    '                'edit it?
    '                MessageBox.Show("Editing Assembly: " + occ.Name)
    '                occ.Edit()
    '                Exit For
    '            End If
    '        Next
    '        Try
    '            If TypeOf occ.Definition.Document Is AssemblyDocument Then
    '                Dim oassy As AssemblyDocument
    '                oassy = occ.Definition.Document
    '                realOcc = oassy.ComponentDefinition.Occurrences.Add(newfilename, PosnMatrix)
    '            Else
    '                realOcc = occ.ContextDefinition.Occurrences.Add(newfilename, PosnMatrix)
    '            End If
    '            realOccStr = realOcc.Name
    '        Catch ex As Exception
    '            MessageBox.Show("Exception was: " + ex.Message + vbCrLf + ex.StackTrace)
    '        End Try
    '    End If

    '    'Assign iProperties
    '    iProperties.Value(realOccStr, "Project", "Description") = subObject.LegacyDescr
    '    iProperties.Value(realOccStr, "Project", "Part Number") = subObject.Partno
    '    iProperties.Value(realOccStr, "Project", "Revision Number") = subObject.LegacyRev
    '    iProperties.Value(realOccStr, "Summary", "Subject") = subobject.LegacyDrawingno
    '    iProperties.Value(realOccStr, "Summary", "Title") = subobject.LegacyDescr
    '    iProperties.Value(realOccStr, "Summary", "Comments") = "MODELLED FROM DRAWINGS"
    '    iProperties.Value(realOccStr, "Project", "Project") = "A90.1"
    '    Try
    '        'occ.ExitEdit(ExitTypeEnum.kExitToParent)
    '    Catch ex As Exception
    '        'occ wasn't activated for editing.
    '    End Try
    'End Sub

    'Private Function FindFileInVWF(newfilename As String) As String
    '    Dim dir = New DirectoryInfo("C:\Vault Working Folder\Designs")
    '    Dim tmpstr = GetExistingFile(dir, System.IO.Path.GetFileNameWithoutExtension(newfilename))
    '    If tmpstr = "" Then
    '        Return ""
    '    Else
    '        Return ""
    '    End If
    'End Function

    'Private Function GetExistingFile(ByVal dir As DirectoryInfo, ByVal newfilename As String) As String
    '    Dim foundfilename As String = String.Empty
    '    For Each file As FileInfo In dir.GetFiles()
    '        If System.IO.Path.GetFileNameWithoutExtension(foundfilename) = newfilename Then
    '            foundfilename = newfilename
    '            Exit For
    '        End If
    '    Next
    '    For Each subDir As DirectoryInfo In dir.GetDirectories()
    '        foundfilename = GetExistingFile(subDir, newfilename)
    '    Next
    '    Return foundfilename
    'End Function

    'Public Class SubObjectCls
    '    Implements IComparable(Of SubObjectCls)
    '    Public PartNo As String
    '    Public LegacyDescr As String
    '    Public LegacyRev As String
    '    Public LegacyDrawingNo As String
    '    Public ParentAssembly As String
    '    Public HasChildren As Boolean
    '    Public Children As List(Of SubObjectCls)

    '    Public Sub Init(m_partno As String,
    '                    m_legacydescr As String,
    '                    m_legacyrev As String,
    '                    m_legacydrawingno As String,
    '                    m_parentassy As String,
    '                    m_haschildren As Boolean,
    '                    m_children As List(Of SubObjectCls))
    '        PartNo = m_partno
    '        LegacyDescr = m_legacydescr
    '        LegacyRev = m_legacyrev
    '        LegacyDrawingNo = m_legacydrawingno
    '        ParentAssembly = m_parentassy
    '        HasChildren = m_haschildren
    '        Children = m_children
    '    End Sub
    '    Public Function CompareTo(other As SubObjectCls) As Integer Implements IComparable(Of SubObjectCls).CompareTo
    '        Return Me.CompareTo(other)
    '    End Function
    'End Class
End Module