Attribute VB_Name = "Module1"
Sub ROSI_Confirmation_Registration()

Dim browser As Object
Dim Workbook_Name As String
Dim Worksheet_Name As String
Dim studentid As Range
Dim Fall_Session As String
Dim Winter_Session As String
Dim ROSI_Name As String
Dim ROSI_Start_Session As String
Dim ROSI_Year_Of_Study As String
Dim Start_Session As String
Dim ROSI_POSt_Cd As String
Dim ROSI_Status As String
Dim ROSI_Reg_20199 As String
Dim ROSI_Reg_20201 As String
Dim Graduation_Date As String
Dim Student_Number As String
Dim msword As Object


Workbook_Name = ActiveWorkbook.Name
Worksheet_Name = ActiveSheet.Name

Set studentid = Application.InputBox("Click the first student number, and drag down to the last student number to highlight all the student numbers. Do not include the header.", _
Title:="Select Range of Student Numbers", _
Type:=8)

Set browser = CreateObject("internetexplorer.application")
Set msword = CreateObject("Word.Application")

With msword
        .Visible = True
        .Documents.Open "C:\Users\olive\OneDrive - University of Toronto\VBA Instructions\Letter of Registration\Template.docx"
        .Activate
        .Visible = False
End With

browser.Navigate "https://javnat-admin-qa.easi.utoronto.ca/ROSI/rosi"
browser.Visible = True



On Error Resume Next
browser.document.all("JavNat_PF12").Click

While browser.Busy
DoEvents
Wend

browser.document.all("field2").Value = "4 A C B"
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

For Each Cell In studentid

Cell.Activate

Fall_Session = "20199"
Winter_Session = "20201"

browser.document.all("#ACTION").Value = "C"

While browser.Busy
DoEvents
Wend

browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

browser.document.all("#ACTION").Value = "D"
browser.document.all("SRA506.SESSION_CD").Value = Fall_Session
browser.document.all("SRA506.PERSON_ID").Value = Cell.Value
browser.document.all("SRA506.POST_CD").Value = ""
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

For Each jtag In browser.document.getelementsbytagname("div")
    If jtag.getattribute("id") = "Previous_Screen_1" Then
        ROSI_POSt_Cd = browser.document.all("line7").innerhtml
        Cells(ActiveCell.Row, 19).Value = ROSI_POSt_Cd
        browser.document.all("JavNat_PF2").Click
        While browser.Busy
        DoEvents
        Wend
        Application.Wait (Now + TimeValue("0:00:02"))
        browser.document.all("SRA506.POST_CD").Value = ROSI_POSt_Cd
        browser.document.all("#ACTION").Value = "D"
        browser.document.all("JavNat_ENTR").Click
        Cells(ActiveCell.Row, 21).Value = "Multiple POSts exist"
        
        Exit For
End If
Next jtag

While browser.Busy
DoEvents
Wend

ROSI_Name = browser.document.all("#DISPLAY-NAME").innerhtml
ROSI_Start_Session = browser.document.all("SRA506.CANDIDACY_SESS_CD").innerhtml
ROSI_Year_Of_Study = browser.document.all("SRA506.YEAR_OF_STUDY").Value
ROSI_POSt_Cd = browser.document.all("SRA506.CANDIDACY_POST_CD").innerhtml
ROSI_Status = browser.document.all("SRA506.ATTENDANCE_CLASS").Value
ROSI_Reg_20199 = browser.document.all("SRA506.CURR_REG_STS_CD").Value

'if REG or INVIT for 20199:

If ROSI_Reg_20199 = "REG" Or ROSI_Reg_20199 = "INVIT" Then
Cells(ActiveCell.Row, 14).Value = Trim(ROSI_Name)
Cells(ActiveCell.Row, 15).Value = ROSI_Reg_20199
Cells(ActiveCell.Row, 17).Value = ROSI_Start_Session
Cells(ActiveCell.Row, 18).Value = ROSI_Year_Of_Study
Cells(ActiveCell.Row, 19).Value = Trim(ROSI_POSt_Cd)
Cells(ActiveCell.Row, 20).Value = ROSI_Status

browser.document.all("#ACTION").Value = "C"
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

browser.document.all("#ACTION").Value = "D"
browser.document.all("SRA506.SESSION_CD").Value = Winter_Session
browser.document.all("SRA506.PERSON_ID").Value = Cell.Value
browser.document.all("SGA506.POST_CD").Value = ""
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

For Each jtag In browser.document.getelementsbytagname("div")
    If jtag.getattribute("id") = "Previous_Screen_1" Then
        ROSI_POSt_Cd = browser.document.all("line7").innerhtml
        Cells(ActiveCell.Row, 19).Value = ROSI_POSt_Cd
        browser.document.all("JavNat_PF2").Click
        While browser.Busy
        DoEvents
        Wend
        Application.Wait (Now + TimeValue("0:00:02"))
        browser.document.all("SRA506.POST_CD").Value = ROSI_POSt_Cd
        browser.document.all("#ACTION").Value = "D"
        browser.document.all("JavNat_ENTR").Click
        Cells(ActiveCell.Row, 21).Value = "Multiple POSts exist"
        
        Exit For
End If
Next jtag

While browser.Busy
DoEvents
Wend

'if Fall and Winter POSTs differ:
If Trim(browser.document.all("SRA506.CANDIDACY_POST_CD").innerhtml) = Cells(ActiveCell.Row, 19).Value Then

ROSI_Reg_20201 = browser.document.all("SRA506.CURR_REG_STS_CD").Value
Cells(ActiveCell.Row, 16).Value = ROSI_Reg_20201
Else: Cells(ActiveCell.Row, 21).Value = "Winter Session POSt differs from Fall Session"
End If

Else: browser.document.all("#ACTION").Value = "C"

While browser.Busy
DoEvents
Wend

browser.document.all("#ACTION").Value = "D"
browser.document.all("SRA506.SESSION_CD").Value = Winter_Session
browser.document.all("SRA506.PERSON_ID").Value = Cell.Value
browser.document.all("SGA506.POST_CD").Value = ""
browser.document.all("JavNat_ENTR").Click

While browser.Busy
DoEvents
Wend

For Each jtag In browser.document.getelementsbytagname("div")
    If jtag.getattribute("id") = "Previous_Screen_1" Then
        ROSI_POSt_Cd = browser.document.all("line7").innerhtml
        Cells(ActiveCell.Row, 19).Value = ROSI_POSt_Cd
        browser.document.all("JavNat_PF2").Click
        While browser.Busy
        DoEvents
        Wend
        Application.Wait (Now + TimeValue("0:00:02"))
        browser.document.all("SRA506.POST_CD").Value = ROSI_POSt_Cd
        browser.document.all("#ACTION").Value = "D"
        browser.document.all("JavNat_ENTR").Click
        Cells(ActiveCell.Row, 21).Value = "Multiple POSts exist"
        
        Exit For
End If
Next jtag

While browser.Busy
DoEvents
Wend

ROSI_Name = browser.document.all("#DISPLAY-NAME").innerhtml
ROSI_Start_Session = browser.document.all("SRA506.CANDIDACY_SESS_CD").innerhtml
ROSI_Year_Of_Study = browser.document.all("SRA506.YEAR_OF_STUDY").Value
ROSI_POSt_Cd = browser.document.all("SRA506.CANDIDACY_POST_CD").innerhtml
ROSI_Status = browser.document.all("SRA506.ATTENDANCE_CLASS").Value
ROSI_Reg_20201 = browser.document.all("SRA506.CURR_REG_STS_CD").Value


Cells(ActiveCell.Row, 15).Value = "Not REG"

If ROSI_Reg_20201 <> "REG" And ROSI_Reg_20201 <> INVIT Then
Cells(ActiveCell.Row, 16).Value = "Not REG"

Else:
Cells(ActiveCell.Row, 14).Value = Trim(ROSI_Name)
Cells(ActiveCell.Row, 16).Value = ROSI_Reg_20201
Cells(ActiveCell.Row, 17).Value = ROSI_Start_Session
Cells(ActiveCell.Row, 18).Value = ROSI_Year_Of_Study
Cells(ActiveCell.Row, 19).Value = Trim(ROSI_POSt_Cd)
Cells(ActiveCell.Row, 20).Value = ROSI_Status

End If

End If

Cells(ActiveCell.Row, 22).Value = Application.VLookup(Cells(ActiveCell.Row, 19), Range("POST_INFO"), 5, False)
Cells(ActiveCell.Row, 23).Value = Application.VLookup(Cells(ActiveCell.Row, 19), Range("POST_INFO"), 6, False)

If Cells(ActiveCell.Row, 20).Value = "FT" Then
Cells(ActiveCell.Row, 20).Value = "Full Time"
End If

If Mid(Cells(ActiveCell.Row, 17), 5, 5) = "1" Then
Start_Session = "January 1, " & Mid(Cells(ActiveCell.Row, 17), 1, 4)
End If

If Mid(Cells(ActiveCell.Row, 17), 5, 5) = "5" Then
Start_Session = "May 1, " & Mid(Cells(ActiveCell.Row, 17), 1, 4)
End If

If Mid(Cells(ActiveCell.Row, 17), 5, 5) = "9" Then
Start_Session = "September 1, " & Mid(Cells(ActiveCell.Row, 17), 1, 4)
End If

Graduation_Date = Format(Cells(ActiveCell.Row, 12), "Long Date")

'If Cells(ActiveCell.Row, 15).Value = "REG" And Cells(ActiveCell.Row, 16).Value = "REG" Then

    'Set ws = ActiveSheet
    'Set msword = CreateObject("Word.Application")

'Set msword = CreateObject("Word.Application")

If Cells(ActiveCell.Row, 15).Value <> "Not REG" And Cells(ActiveCell.Row, 16).Value <> "Not REG" Then
With msword
        .Visible = False
        '.Documents.Open "C:\Users\olive\OneDrive - University of Toronto\VBA Instructions\Letter of Registration\Template.docx"
        .Activate
        
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[ROSI_Name]"
            .Replacement.Text = Cells(ActiveCell.Row, 14).Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
         With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Transcript_Title]"
            .Replacement.Text = Cells(ActiveCell.Row, 23).Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
'do INVIT have year of study?

        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[ROSI_Year_Of_Study]"
            .Replacement.Text = Cells(ActiveCell.Row, 18).Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Unit_Name]"
            .Replacement.Text = Cells(ActiveCell.Row, 22).Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
  
'do INVIT have ROSI status? (FT/PT?)
  
  If Cells(ActiveCell.Row, 10).Value = "Yes" Then
  
          With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[ROSI_Status]"
            .Replacement.Text = Cells(ActiveCell.Row, 20).Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
    

    Else:
            With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[ROSI_Status]"
            .Replacement.Text = ""

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
            With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "Status:"
            .Replacement.Text = ""

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
    End If
    
    'do INVIT have start session?
    
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Start_Date]"
            .Replacement.Text = Start_Session

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Student_Number]"
            .Replacement.Text = ActiveCell.Value

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
  If Cells(ActiveCell.Row, 11).Value = "Yes" And Cells(ActiveCell.Row, 12) <> "" Then

        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Graduation_Date]"
            .Replacement.Text = Graduation_Date

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
        Else:
        
            With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = ", and is expected to complete the degree requirements by [Graduation_Date]."
            .Replacement.Text = "."

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
        
        
        
    End If
        
    If Cells(ActiveCell.Row, 9).Value = "Full academic year (2019-2020 Fall/Winter)" And Cells(ActiveCell.Row, 15).Value = "REG" And Cells(ActiveCell.Row, 16).Value = "REG" Then
  
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Academic_Session]"
            .Replacement.Text = "2019-2020 Fall/Winter session (September 1, 2019 - April 30, 2020)"

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
  End If
  
      If Cells(ActiveCell.Row, 9).Value = "Full academic year (2019-2020 Fall/Winter)" And Cells(ActiveCell.Row, 15).Value <> "REG" And Cells(ActiveCell.Row, 16).Value = "REG" Then
  
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Academic_Session]"
            .Replacement.Text = "2020 Winter session (January 1, 2020 - April 30, 2020)"

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
  End If
  
      If Cells(ActiveCell.Row, 9).Value = "Full academic year (2019-2020 Fall/Winter)" And Cells(ActiveCell.Row, 15).Value = "REG" And Cells(ActiveCell.Row, 16).Value <> "REG" Then
  
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Academic_Session]"
            .Replacement.Text = "2019 Fall session (September 1, 2019 - December 31, 2019)"

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
  End If
  
        
    If Cells(ActiveCell.Row, 9).Value = "Current Session (2020 Winter)" And Cells(ActiveCell.Row, 16).Value = "REG" Then
  
        With .ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            .Text = "[Academic_Session]"
            .Replacement.Text = "2020 Winter session (January 1, 2020 - April 30, 2020)"

            .Forward = True
            .Wrap = 1               'wdFindContinue (WdFindWrap Enumeration)
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False

            .Execute Replace:=2     'wdReplaceAll (WdReplace Enumeration)
        End With
  End If
    
        
     End With
     
  
    msword.ActiveDocument.ExportAsFixedFormat OutputFileName:="C:\Users\olive\OneDrive - University of Toronto\VBA Instructions\Letter of Registration\Letters\" & Trim(ROSI_Name) & " - Confirmation of Registration.pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
        wdExportDocumentContent, IncludeDocProps:=False, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False


msword.ActiveDocument.Undo (15)
msword.ActiveDocument.Visible = False



'msword.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
'ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
'msword.Visible = False
'msword.Quit
'Application.Quit

End If

Next

msword.ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
msword.Quit

'better incorporate INVIT status for letters for those that are "eligible to enroll" in programs
'change from QA to live version of ROSI
'figure out how to active active IE browser instead of opening new session (doesn't work for live ROSI)


MsgBox ("ROSI registration download and Confirmation of Registration Letters complete!")


End Sub



