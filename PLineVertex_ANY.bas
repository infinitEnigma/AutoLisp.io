Attribute VB_Name = "PLineVertex_ANY"
Sub Vertex_Extract3dPoly()

Dim i As Long
Dim pLine As Variant
Dim pCoord1 As Variant
Dim Pt(2) As Double
Dim txtVertex As AcadText
    ' priprema za unos podataka u txt fajl
    Set fs = CreateObject("Scripting.FileSystemObject")
    ' prolazi kroz sve objekte u aktivnom layout-u
    For a = 0 To (ThisDrawing.ActiveLayout.Block.Count - 1)
        ' trazi 3d poliliniju
        If ThisDrawing.ActiveLayout.Block.Item(a).ObjectName = "AcDb3dPolyline" Then
            
            ' otvara novi txt fajl i dodaje index linije kao parametar imena fajla
            Set b = fs.CreateTextFile("E:\geodux\3d-poli" & a & ".txt", True)
            Set pLine = ThisDrawing.ModelSpace.Item(a)
            pCoord1 = pLine.Coordinates
            Pt(0) = pCoord1(0)
            Pt(1) = pCoord1(1)
            Pt(2) = pCoord1(2)
            icount = UBound(pCoord1)
            ' prolazi kroz sve tacke na 3d poliliniji
            For i = 0 To icount - 1 Step 3
        
                On Error Resume Next
                ' upisuje koordinate u txt fajl
                b.write ("pr" & i / 3 + 1) & vbTab
                b.write (Format(pCoord1(i), "#.##")) & vbTab
                b.write (Format(pCoord1(i + 1), "#.##")) & vbTab
                b.write (Format(pCoord1(i + 2), "#.##")) & vbCrLf
            
            
                Pt(0) = pCoord1(i)
                Pt(1) = pCoord1(i + 1)
                Pt(2) = pCoord1(i + 2)
                ' upisuje naziv tacke na 3d poliliniji u crtezu
                Set txtVertex = ThisDrawing.ModelSpace.AddText("pr" & i / 3 + 1, Pt, 5)
            
        
            Next i
        ElseIf ThisDrawing.ActiveLayout.Block.Item(a).ObjectName = "AcadLWPolyline" Or ThisDrawing.ActiveLayout.Block.Item(a).ObjectName = "AcDbPolyline" Then
            ' otvara novi txt fajl i dodaje index linije kao parametar imena fajla
            Set b = fs.CreateTextFile("E:\geodux\2d-poli" & a & ".txt", True)
            Set pLine = ThisDrawing.ModelSpace.Item(a)
            pCoord1 = pLine.Coordinates
            Pt(0) = pCoord1(0)
            Pt(1) = pCoord1(1)
            icount = UBound(pCoord1)
            ' prolazi kroz sve tacke na 2d poliliniji
            For i = 0 To icount - 1 Step 2
        
                On Error Resume Next
                ' upisuje koordinate u txt fajl
                b.write ("pr" & i / 2 + 1) & vbTab
                b.write (Format(pCoord1(i), "#.##")) & vbTab
                b.write (Format(pCoord1(i + 1), "#.##")) & vbTab
                
            
                Pt(0) = pCoord1(i)
                Pt(1) = pCoord1(i + 1)
                ' upisuje naziv tacke na 2d poliliniji u crtezu
                Set txtVertex = ThisDrawing.ModelSpace.AddText("pr" & i / 2 + 1, Pt, 5)
            Next i
          End If
    Next a

End Sub
