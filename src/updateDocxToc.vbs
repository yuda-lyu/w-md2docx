' updateDocxToc.vbs: update fields of a docx with local Microsoft Word and save as another docx (used by addDocxToc.mjs)
' SEQ captions, table of contents and tables of figures need pagination that only Word can compute;
' entries are rebuilt first, then the document is repaginated and page numbers are recomputed (the TOC itself changes the page count)
' Keep this file ASCII only: cscript reads .vbs in the system code page, non-ASCII string literals break compilation
' usage: cscript //nologo updateDocxToc.vbs <src docx> <dst docx>   (absolute paths)
' stdout: "toc=<n> tof=<n> pages=<n>" then "ok" on success; "error: <reason>" and exit code 1 on failure
Option Explicit
Dim args, word, doc, t, f, r
Set args = WScript.Arguments
If args.Count < 2 Then
    WScript.Echo "error: usage: cscript //nologo updateDocxToc.vbs <src docx> <dst docx>"
    WScript.Quit 2
End If
On Error Resume Next
Set word = CreateObject("Word.Application")
If Err.Number <> 0 Then
    WScript.Echo "error: cannot start Microsoft Word: " & Err.Description
    WScript.Quit 1
End If
word.Visible = False
word.DisplayAlerts = 0
Set doc = word.Documents.Open(args(0), False, False, False)
If Err.Number <> 0 Then
    WScript.Echo "error: open failed: " & Err.Description
    word.Quit 0
    WScript.Quit 1
End If
r = doc.Fields.Update
doc.Repaginate
For Each t In doc.TablesOfContents
    t.Update
Next
For Each f In doc.TablesOfFigures
    f.Update
Next
doc.Repaginate
For Each t In doc.TablesOfContents
    t.UpdatePageNumbers
Next
For Each f In doc.TablesOfFigures
    f.UpdatePageNumbers
Next
If Err.Number <> 0 Then
    WScript.Echo "error: update failed: " & Err.Description
    doc.Close 0
    word.Quit 0
    WScript.Quit 1
End If
WScript.Echo "toc=" & doc.TablesOfContents.Count & " tof=" & doc.TablesOfFigures.Count & " pages=" & doc.ComputeStatistics(2)
doc.SaveAs2 args(1), 12
If Err.Number <> 0 Then
    WScript.Echo "error: save failed: " & Err.Description
    doc.Close 0
    word.Quit 0
    WScript.Quit 1
End If
doc.Close 0
word.Quit 0
WScript.Echo "ok"
