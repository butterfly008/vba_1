Attribute VB_Name = "Module3"
Sub getnum_v2()
Dim stlen As String
Dim i As Integer
Dim arra() As String
Dim arran() As String


Orig.AutoFilterMode = False
Call BeginMacro

LastRow = Orig.Cells(Rows.Count, 1).End(xlUp).Row
Orig.Range("J2:J" & LastRow).clear



For n = 1 To LastRow
celref = Orig.Cells(n, 4).Value


arra() = Split(celref, " ")

For counter = LBound(arra) To UBound(arra)

  strin = arra(counter)
  
  storage = Trim(strin)
   
   lenof = Len(storage)
               
                If lenof = 9 Then
                
              somstr = Mid(storage, 1, 1)
                  somot = Mid(storage, 9, 1)
                 
                
                                If somstr = "P" Or somstr = "p" And IsNumeric(somot) = True Then
                              
                   
                                    storage = Right(storage, 7)
                   
                                    Orig.Cells(n, 10).Value = storage
                                  
                                    End If
                   
                         ElseIf lenof = 10 Then
                  
                            somstr = Mid(storage, 1, 1)
                   somot = Mid(storage, 10, 1)
                
                            If somstr = "P" Or somstr = "p" And IsNumeric(somot) = True Then
                             
                            storage = Right(storage, 7)
                            Orig.Cells(n, 10).Value = storage
                          
                            End If
                            
                         ElseIf lenof = 8 Then
                                            somstr = Mid(storage, 1, 1)
                                       somot = Mid(storage, 8, 1)
                               If somstr = "-" Or somstr = "#" And IsNumeric(somot) = True Then
                                   
                   
                                    storage = Right(storage, 7)
                   
                                    Orig.Cells(n, 10).Value = storage
                                  
                                    End If
         
                        End If
                  
                arran() = Split(storage, ",")
                
              If Orig.Cells(n, 10).Value <> storage Then
              
                    For counter2 = LBound(arran) To UBound(arran)
           
                        strin2 = arran(counter2)
  
                        storage2 = Trim(strin2)
  
                If IsNumeric(storage2) = True And Len(storage2) = 7 Then
                        car = Mid(storage2, 1, 1)
                        
                        If car = 1 Then
    
                        Orig.Cells(n, 10).Value = storage2
    
  'Orig.Cells(n, 10).Value = "no po# in D"
                        End If
                  Else
                        
                       If IsNumeric(Orig.Cells(n, 10).Value) = True And Len(Orig.Cells(n, 10).Value) = 7 Then
                         
                               Orig.Cells(n, 10).Value = Orig.Cells(n, 10).Value
                        
                       Else
                              
                         Orig.Cells(n, 10).Value = "no po# in D"
                         
                      End If
                      
                       
               End If
  
                Next counter2
            End If
   
  Next counter
  
  Next n

Call EndMacro


End Sub
