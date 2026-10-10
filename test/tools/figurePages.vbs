' figurePages.vbs
' Copied from w-html2docx 1.0.34 test/tools/figurePages.vbs (a test asset of that package, not part of its API).
' Usage: cscript //nologo figurePages.vbs <docx>
' Opens the docx read-only in an invisible Microsoft Word. For each paragraph holding an inline picture, prints one line
'   <picture page>,<caption first line page>,<caption last line page>
' where the caption is the next paragraph whose text is not empty after removing the paragraph mark, U+200B and spaces,
' and prints 0 for both caption pages when there is none. Word is always closed; exits with 1 on error.
' ASCII only: cscript reads the script with the system code page.
Option Explicit


Function PageOf(d, pos)
    PageOf = d.Range(pos, pos).Information(3) 'wdActiveEndPageNumber
End Function


Sub ListPages(d)
    Dim n, i, j, p, q, t, pg, cs, ce
    n = d.Paragraphs.Count
    For i = 1 To n
        Set p = d.Paragraphs(i)
        If p.Range.InlineShapes.Count > 0 Then
            pg = PageOf(d, p.Range.Start)
            cs = 0
            ce = 0
            For j = i + 1 To n
                Set q = d.Paragraphs(j)
                t = Replace(Replace(Replace(q.Range.Text, vbCr, ""), ChrW(8203), ""), " ", "")
                If Len(t) > 0 Then
                    cs = PageOf(d, q.Range.Start)
                    ce = PageOf(d, q.Range.End - 1)
                    Exit For
                End If
            Next
            WScript.Echo pg & "," & cs & "," & ce
        End If
    Next
End Sub


Dim fp, app, doc, rc
rc = 0
fp = WScript.Arguments(0)
Set app = CreateObject("Word.Application")
app.Visible = False
app.DisplayAlerts = 0
On Error Resume Next
Set doc = app.Documents.Open(fp, False, True, False)
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine "error: " & Err.Description
    rc = 1
Else
    doc.Repaginate
    ListPages doc
    If Err.Number <> 0 Then
        WScript.StdErr.WriteLine "error: " & Err.Description
        rc = 1
    End If
    doc.Close 0
End If
app.Quit 0
WScript.Quit rc
