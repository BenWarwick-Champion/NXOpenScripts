Option Strict Off
Imports System
Imports System.Collections.Generic
Imports NXOpen
Imports NXOpen.UF

Module reportMissingBalloons

    Dim theSession As Session = Session.GetSession()
    Dim theUfSession As UFSession = UFSession.GetUFSession()

    Dim lw As ListingWindow = theSession.ListingWindow

    Sub Main()

        If IsNothing(theSession.Parts.Work) Then
            'Active Part Required
            Return
        End If

        Dim workPart As Part = theSession.Parts.Work
        lw.Open()

        Dim partListTags() As Tag = IsNothing
        Dim numPartsLists As Integer = 0
        theUfSession.Plist.AskTags(partListTags, numPartsLists)

        'Test for a dummy parts list
        If numPartsLists > 0 Then
            For i As Integer = 0 To partListTags.Length - 1
                Dim prtList As DisplayableObject = Utilities.NXObjectManager.Get(partListTags(i))
                If prtList.Layer = 0 Then
                    numPartsLists -= 1
                End If
            Next
        End If

        If numPartsLists <= 0 Then
            lw.WriteLine("No parts list found in the work part")
            Return
        End If

        If numPartsLists > 1 Then
            lw.WriteLine("ERROR: Cannot evaluate AutoBalloons with multiple parts lists")
            lw.WriteLine("Check environment variable UGII_UPDATE_ALL_ID_SYMBOLS_WITH_PLIST")
            Return
        End If

        theUfSession.Plist.UpdateAllPlists()

        Dim numRows As Integer
        theUfSession.Tabnot.AskNmRows(partListTags(0), numRows)

        Dim plistPrefs As UFPlist.Prefs = Nothing
        theUfSession.Plist.AskPrefs(partListTags(0), plistPrefs)

        Dim colCallout As Tag = PartListCalloutColumn(partListTags(0))
        If colCallout = Tag.Null Then
            lw.WriteLine("Parts list callout column not found")
            Return
        End If

        Dim plBalloonList As New List(Of String)
        CollectPartsListBalloons(plBalloonList, plistPrefs.symbol_type)

        Dim missingBalloons As New List(Of String)

        Dim startTime As DateTime = Now

        'Loop through rows of the parts list. Look for corresponding callout.
        For i As Integer = 0 To numRows - 1

            Dim rowTag As Tag
            theUfSession.Tabnot.AskNthRow(partListTags(0), i, rowTag)
            Dim cellTag As Tag
            theUfSession.Tabnot.AskCellAtRowCol(rowTag, colCallout, cellTag)
            Dim calloutEvText As String = ""
            theUfSession.Tabnot.AskEvaluatedCellText(cellTag, calloutEvText)

            If plBalloonList.Contains(calloutEvText) Then
                'Balloon found. Remove it from the list.
                plBalloonList.Remove(calloutEvText)
            Else
                'Balloon NOT found. Add to "missing" list.
                missingBalloons.Add(calloutEvText)
            End If

        Next

        Dim endTime As DateTime = Now

        If missingBalloons.Count = 0 Then
            lw.WriteLine("All parts in the BOM have a corresponding callout")
        Else If missingBalloons.Count = 1 Then
            lw.WriteLine(missingBalloons.Count.ToString & " part(s) in the BOM is missing a callout")
            For Each missing As String In missingBalloons
                lw.WriteLine("Item: " & missing)
            Next
        End If

        lw.Close()

    End Sub

    Function PartListCalloutColumn(ByVal partListTag As Tag) As Tag
        Dim numColumns As Integer
        theUfSession.Tabnot.AskNmColumns(partListTag, numColumns)
        Dim rowTag As Tag
        theUfSession.Tabnot.AskNthRow(partListTag, 0, rowTag)

        For j As Integer = 0 To numColumns - 1
            Dim colTag As Tag
            theUfSession.Tabnot.AskNthColumn(partListTag, j, colTag)
            Dim cellTag As Tag
            theUfSession.Tabnot.AskCellAtRowCol(rowTag, colTag, cellTag)

            'Get the current cell text
            Dim cellText As String = ""
            theUfSession.Tabnot.AskCellText(cellTag, cellText)
            If cellText = "$~C" Then
                Return colTag

            End If

        Next

        Return Nothing

    End Function

    Sub CollectPartsListBalloons(ByRef theBalloonList As List(Of String), ByVal plSymbolType As Integer)
        For Each tempId As Annotations. IdSymbol In theSession.Parts.Work.Annotations.IdSymbols
            Dim theIdSymbolBuilder As Annotations.IdSymbolBuilder = theSession.Parts.Work.Annotations.IdSymbols.CreateIdSymbolBuilder(tempId)

            If plSymbolType = theIdSymbolBuilder.Type + 1 Then
                'Symbol matches type used in parts list
            Else
                'Symbol does not match, skip it
                Continue For
            End If

            If Not theBalloonList.Contains(theIdSymbolBuilder.UpperText) Then
                theBalloonList.Add(theIdSymbolBuilder.UpperText)
            End If

            theIdSymbolBuilder.Destroy()

        Next
    End Sub

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        'Unloads the image immediately after execution within NX
        GetUnloadOption = NXOpen.Session.LibraryUnloadOption.Immediately

    End Function

End Module
