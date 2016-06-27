Attribute VB_Name = "Module1"

Dim sf2pBags As Integer
Dim sf1pBags As Integer
Dim rad1pBags As Integer
Dim pea2pBags As Integer
Dim pea1pBags As Integer
Dim wg1pBags As Integer
Dim wg2pBags As Integer
Dim bw1pBags As Integer

'Seeding
Dim sowDay As Date
Dim sowCrop As String
Dim sowTrays As Double

Dim harvestDate As Date
Dim q As Double  'Quantities in pounds
Dim crop As String
Dim customer As String
Dim route As String
Dim price As String
Dim size As String
Dim oddWeek As Integer
Dim secondPage As String
Dim spacing As Integer
Dim pay As Integer

Dim fcolor As Long
Dim customerCell As String 'Column G + row xx
Dim switchBack As String 'To save a cell address if need to return
Dim labelColumn As Integer 'Column label 1, 2 or 3
Dim picCell As String 'Cell where the Logo need to be inserted

Dim sfSmallBags As Integer
Dim peaSmallBags As Integer
Dim radishSmallBags As Integer
Dim startingCell As Range

Dim sfClick As Boolean 'Used to determine when (after the first large bag) small bags need to be counted
Dim peaClick As Boolean 'Used to determine when (after the first large bag) small bags need to be counted
Dim radishClick As Boolean 'Used to determine when (after the first large bag) small bags need to be counted


'///////////////////////////////////////////////////////////////////////////
'LM is the master SUB that calls on the rest, its the skeleton of the Marco
'///////////////////////////////////////////////////////////////////////////

Sub LM()
Attribute LM.VB_ProcData.VB_Invoke_Func = "l\n14"
    ' Keyboard Shortcut: Ctrl+l
    
    'Set the routes table size (Not in use right now)
    'spacing = 60
    
    ClearSheets
    
    labelColumn = 1
    
    'Add seeding numbers for the week.
     Sheets("Routes").Select
     Range("J23").Select
     Seeding
     Sheets("Routes").Select
     Range("J100").Select
     Seeding
    
    'Start from the top??
     Sheets("LabelMaker").Select
     Range("A1").Select
    
    'Start from the top??
     Sheets("Routes").Select
     Range("A4").Select
     
    Sheets("ExpectedSales").Select
            
    'Get starting position
    startingPosition = ActiveCell.Address
        
    For i = 1 To 2    'Tuesday & Friday
         
        'Reset small, Large bag numbers and clicks
        sfClick = False
        peaClick = False
        radishClick = False
        sfSmallBags = 0
        peaSmallBags = 0
        radishSmallBags = 0
        
        sf1pBags = 0
        sf2pBags = 0
        pea1pBags = 0
        pea2pBags = 0
        rad1pBags = 0
        wg2pBags = 0
        wg1pBags = 0
        bw1pBags = 0
        
        'Get date
        harvestDate = ActiveCell
        
        'Add dates to the Order Sheet
        Sheets("Routes").Select
        If i = 1 Then
            Range("A2").Select
            ActiveCell = harvestDate
            Range("A4").Select
        Else
            Range("A80").Select
            ActiveCell = harvestDate
            Range("A82").Select
        End If
        Sheets("ExpectedSales").Select
        
        'Gotto first entry
        ActiveCell.Offset(7, 0).Select
                
        'This pass will go crop by crop
        
        SunflowerCycle
        
        PeaCycle
        
        RadishCycle
        
        BuckwheatCycle
        
        WheatGrassTraysCycle
        
        WheaGrassBagsCycle
        
        'Create a label that sumarized the amount of small bags to be made
        Sheets("LabelMaker").Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell = "X.X.X.X"
        ActiveCell.Offset(-3, 1).Select
        ActiveCell = "SF: " & sfSmallBags
        ActiveCell.Offset(1, 0).Select
        ActiveCell = "X"
        ActiveCell.Offset(0, 1).Select
        ActiveCell = "X.X"
        ActiveCell.Offset(0, 1).Select
        ActiveCell = "Pea: " & peaSmallBags
        ActiveCell.Offset(1, 0).Select
        ActiveCell = "Rad: " & radishSmallBags
        ActiveCell.Offset(1, 0).Select
        ActiveCell = harvestDate
        Move2next
        
        ' How do you want to space every day? In what Row do the friday routes start?
        Sheets("Routes").Select
        Range("A68").Select
        
        Sheets("ExpectedSales").Select
        ActiveCell.Offset(22, -9).Select
        
        Sheets("Routes").Select
        Select Case i
            Case 1
                Range("K4").Select
            Case 2
                Range("K82").Select
        End Select
        'Add a summary bag count
        AddBagSummay
        
    Next i

    SortRoutes  'By route and by customer
        
    'Go back to starting positions
    Sheets("ExpectedSales").Select
    Range(startingPosition).Select
    Sheets("LabelMaker").Select
    Range("A1").Select
End Sub

'///////////////////////////////////////////////////////////////////////////
'/////////////////////////////////End of LM/////////////////////////////////
'///////////////////////////////////////////////////////////////////////////


'//////////////////////Crop cycles//////////////////////

Sub SunflowerCycle()
' / / / / SUNFLOWER CYCLE STARTS HERE
            crop = "Sunflower Shoots"
            Do While ActiveCell <> "x"
                If ActiveCell = "S" Then
                    size = "Small"
                    ActiveCell.Offset(0, 1).Select
                    If sfClick = True Then
                        If ActiveCell <> "" Then
                            If ActiveCell > 0 Then
                                q = ActiveCell
                                'Who is the customer for this row?
                                GetCustomer
                                GetRoute
                                GetPrice
                                ' Adds armi's basil to the routes even though it not in the LBP
                                'If customer = "Dr. Armitage" Then
                                    'Decides if its an odd or an even week
                                    'oddWeek = Format(harvestDate, "ww") - 2 * Int(Format(harvestDate, "ww") / 2)
                                    'Only adds to route every other week
                                    'If oddWeek = 0 Then
                                        'crop = "Basil"
                                        'q = 2
                                        'AddOrder
                                        'Sheets("ExpectedSales").Select
                                        'crop = "Sunflower Shoots"
                                        'q = ActiveCell
                                    'End If
                                'End If
                                'Create an order in the "routes" sheet
                                 AddOrder
                                 Sheets("ExpectedSales").Select
                                'CSA 80g bags should not be included in the Small bag count as the require a diferent bag & label
                                If customer <> "Harvest(CSA)" Then
                                    sfSmallBags = sfSmallBags + q
                                End If
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "T" Then
                    size = "Tray"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0 Then
                            q = ActiveCell
                            'Who is the customer for this row?
                            GetCustomer
                            GetRoute
                            GetPrice
                            'Create an order in the "routes" sheet
                            AddOrder
                            Sheets("ExpectedSales").Select
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "L" Then
                    sfClick = True
                    size = "Large"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                            If ActiveCell > 0.1 Then
                                GetCustomer
                                If customer <> "BUFFER" Then
                                    DoSomething 'We finally found an order.. now get going!
                                End If
                            End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
            'END OF SUNFLOWER CYCLE
End Sub
Sub PeaCycle()
'START OF PEA CYCLE

            ActiveCell.Offset(0, 2).Select
            Do While ActiveCell <> ""
                ActiveCell.Offset(-1, 0).Select
            Loop
            ActiveCell.Offset(1, 0).Select
            crop = "Pea Shoots"
            Do While ActiveCell <> "x"
                
                If ActiveCell = "S" Then
                    size = "Small"
                    ActiveCell.Offset(0, 1).Select
                    If peaClick = True Then
                        If ActiveCell <> "" Then
                            If ActiveCell > 0 Then
                                q = ActiveCell
                                'Who is the customer for this row?
                                GetCustomer
                                GetRoute
                                GetPrice
                                'Create an order in the "routes" sheet
                                 AddOrder
                                 Sheets("ExpectedSales").Select
                                 'CSA 80g bags should not be included in the Small bag count as the require a diferent bag & label
                                 If customer <> "Harvest(CSA)" Then
                                peaSmallBags = peaSmallBags + q
                                End If
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "T" Then
                    size = "Tray"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0 Then
                            q = ActiveCell
                            'Who is the customer for this row?
                            GetCustomer
                            GetRoute
                            GetPrice
                            'Create an order in the "routes" sheet
                            AddOrder
                            Sheets("ExpectedSales").Select
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "L" Then
                    size = "Large"
                    'Gotto q side and look for quantities
                    peaClick = True
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0.1 Then
                            GetCustomer
                            If customer <> "BUFFER" Then
                                DoSomething 'We finally found an order.. now get going!
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
            'END OF PEA CYCLE
End Sub

Sub RadishCycle()
'START OF RADISH CYCLE
            ActiveCell.Offset(0, 2).Select
            Do While ActiveCell <> ""
                ActiveCell.Offset(-1, 0).Select
            Loop
            ActiveCell.Offset(1, 0).Select
            crop = "Radish Shoots"
            Do While ActiveCell <> "x"
            
                If ActiveCell = "S" Then
                    size = "Small"
                    ActiveCell.Offset(0, 1).Select
                    If radishClick = True Then
                        If ActiveCell <> "" Then
                            If ActiveCell > 0 Then
                                q = ActiveCell
                                'Who is the customer for this row?
                                GetCustomer
                                GetRoute
                                GetPrice
                                'Create an order in the "routes" sheet
                                 AddOrder
                                 Sheets("ExpectedSales").Select
                                 'CSA 80g bags should not be included in the Small bag count as the require a diferent bag & label
                                 If customer <> "Harvest(CSA)" Then
                                    radishSmallBags = radishSmallBags + q
                                End If
                            End If
                        End If
                    End If
                    
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "T" Then
                    size = "Tray"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0 Then
                            q = ActiveCell
                            'Who is the customer for this row?
                            GetCustomer
                            GetRoute
                            GetPrice
                            'Create an order in the "routes" sheet
                            AddOrder
                            Sheets("ExpectedSales").Select
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                
                If ActiveCell = "L" Then
                    size = "Large"
                    radishClick = True
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0.1 Then
                            GetCustomer
                            If customer <> "BUFFER" Then
                                DoSomething 'We finally found an order.. now get going!
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
            'END OF RADISH CYCLE
End Sub

Sub BuckwheatCycle()
'START OF BUCKWHEAT CYCLE
            ActiveCell.Offset(0, 2).Select
            
            Do While ActiveCell <> ""
                ActiveCell.Offset(-1, 0).Select
            Loop
            
            ActiveCell.Offset(1, 0).Select
            crop = "Buckwheat Shoots"
            Do While ActiveCell <> "x"
                If ActiveCell = "L" Then
                    size = "Large"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0.1 Then
                            GetCustomer
                            If customer <> "BUFFER" Then
                                DoSomething 'We finally found an order.. now get going!
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                If ActiveCell = "T" Then
                    size = "Tray"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell > 0 Then
                            q = ActiveCell
                            'Who is the customer for this row?
                            GetCustomer
                            GetRoute
                            GetPrice
                            'Create an order in the "routes" sheet
                            AddOrder
                            Sheets("ExpectedSales").Select
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
            
            
            'END OF BUCKWHEAT CYCLE
End Sub

Sub WheaGrassBagsCycle()
'START OF WHEATGRASS CYCLE
            ActiveCell.Offset(0, 2).Select
            
            Do While ActiveCell <> ""
                ActiveCell.Offset(-1, 0).Select
            Loop
            
            ActiveCell.Offset(1, 0).Select
            crop = "Wheatgrass"
            
            Do While ActiveCell <> "x"
                If ActiveCell = "L" Then
                    size = "Large"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell <> "0" Then
                            GetCustomer
                            If customer <> "BUFFER" Then
                                DoSomething 'We finally found an order.. now get going!
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
End Sub

Sub WheatGrassTraysCycle()
'START OF WGTrays CYCLE
            ActiveCell.Offset(0, 2).Select
            
            Do While ActiveCell <> ""
                ActiveCell.Offset(-1, 0).Select
            Loop
            
            ActiveCell.Offset(1, 0).Select
            crop = "Wheatgrass"
            Do While ActiveCell <> "x"
                If ActiveCell = "T" Then
                    size = "Tray"
                    'Gotto q side and look for quantities
                    ActiveCell.Offset(0, 1).Select
                    If ActiveCell <> "" Then
                        If ActiveCell <> "0" Then
                            GetCustomer
                            If customer <> "BUFFER" Then
                                q = ActiveCell
                                'Who is the customer for this row?
                                GetRoute
                                GetPrice
                                'Create an order in the "routes" sheet
                                AddOrder
                                Sheets("ExpectedSales").Select
                            End If
                        End If
                    End If
                    ActiveCell.Offset(0, -1).Select
                End If
                ActiveCell.Offset(1, 0).Select
            Loop
            'END OF WGTrays CYCLE
End Sub

'Sub MixedShootsCycle()
'START OF Mixed CYCLE
 '           ActiveCell.Offset(0, 8).Select
   '
    '        Do While ActiveCell <> ""
     '           ActiveCell.Offset(-1, 0).Select
      '      Loop
       '     ActiveCell.Offset(1, 0).Select
        '    crop = "Mixed Shoots"
         '   Do While ActiveCell <> "x"
          '
           '     If ActiveCell = "S" Then
            '        size = "Small"
'                    ActiveCell.Offset(0, 1).Select
 '                   If radishClick = True Then
  '                      If ActiveCell <> "" Then
   '                         If ActiveCell > 0 Then
    '                            q = ActiveCell
     '                           'Who is the customer for this row?
      '                          GetCustomer
       '                         GetRoute
        '                        GetPrice
         '                       'Create an order in the "routes" sheet
          '                       AddOrder
           '                      Sheets("ExpectedSales").Select
            '                     'CSA 80g bags should not be included in the Small bag count as the require a diferent bag & label
             '                    If customer <> "Harvest(CSA)" Then
              '                      radishSmallBags = radishSmallBags + q
               '                 End If
                '            End If
                 '       End If
                  '  End If
                    
'                    ActiveCell.Offset(0, -1).Select
 '               End If
  '
   '             If ActiveCell = "T" Then
    '                size = "Tray"
                    'Gotto q side and look for quantities
'                    ActiveCell.Offset(0, 1).Select
 '                   If ActiveCell <> "" Then
  '                      If ActiveCell > 0 Then
   '                         q = ActiveCell
    '                        'Who is the customer for this row?
     '                       GetCustomer
      '                      GetRoute
       '                     GetPrice
        '                    'Create an order in the "routes" sheet
         '                   AddOrder
          '                  Sheets("ExpectedSales").Select
           '             End If
            '        End If
'                    ActiveCell.Offset(0, -1).Select
 '               End If
  '
   '             If ActiveCell = "L" Then
'                    size = "Large"
 '                   radishClick = True
  '                  'Gotto q side and look for quantities
   '                 ActiveCell.Offset(0, 1).Select
    '                If ActiveCell <> "" Then
     '                   If ActiveCell <> "0" Then
      '                      GetCustomer
       '                     If customer <> "BUFFER" Then
        '                        DoSomething 'We finally found an order.. now get going!
         '                   End If
          '              End If
           '         End If
            '        ActiveCell.Offset(0, -1).Select
             '   End If
'                ActiveCell.Offset(1, 0).Select
 '           Loop
  '          'END OF CYCLE
'End Sub

'//////////////////////End of crop cycles//////////////////////

Sub GetCustomer()
    switchBack = ActiveCell.Address
    customerCell = "D" & ActiveCell.Row
    Range(customerCell).Select
    customer = ActiveCell
    Range(switchBack).Select
End Sub

Sub GetRoute()
    switchBack = ActiveCell.Address
    customerCell = "C" & ActiveCell.Row
    Range(customerCell).Select
    route = ActiveCell
    Range(switchBack).Select
End Sub

Sub GetPrice()
    switchBack = ActiveCell.Address
    customerCell = "E" & ActiveCell.Row
    Range(customerCell).Select
    price = ActiveCell
    Range(switchBack).Select
End Sub

Sub DoSomething()
    q = ActiveCell
    'Who is the customer for this row?
    GetCustomer
    GetRoute
    GetPrice
    'Create an order in the "routes" sheet
     AddOrder
     
    'Create label(s)
    Sheets("LabelMaker").Select
    
    If crop = "Radish Shoots" Then
         Do While q >= 1
            'insert Logo
            Logo
            'fill label
            Fill1
            'move to next label
            Move2next
        Loop
    Else
        Do While q >= 2
            'insert Logo
            Logo
            'fill label
            Fill2
            'move to next label
            Move2next
        Loop
    End If
        Do While q > 0
            'insert Logo
            Logo
            'fill label
            Fillq
            'move to next label
            Move2next
        Loop
    
    Sheets("ExpectedSales").Select
End Sub

Sub AddOrder()

Sheets("Routes").Select
    Select Case route
        Case "DK"
            fcolor = RGB(220, 10, 10)
        Case "Downtown"
            fcolor = RGB(10, 10, 10)
         Case "East-DT"
           fcolor = RGB(100, 10, 100)
         Case "Drive"
            fcolor = RGB(100, 10, 10)
         Case "Local"
            fcolor = RGB(100, 100, 110)
         Case "Main"
            fcolor = RGB(10, 10, 10)
        Case "NB"
            fcolor = RGB(10, 100, 100)
         Case "Pickup"
            fcolor = RGB(10, 100, 200)
         Case "PNE"
            fcolor = RGB(10, 100, 100)
        Case "West4th"
            fcolor = RGB(150, 10, 100)
        Case "Broadway"
             fcolor = RGB(10, 100, 10)
        Case "Yaletown"
             fcolor = RGB(160, 32, 240)
    End Select
    
     
     ActiveCell.Font.Color = fcolor
     ActiveCell = ""
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = route
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = customer
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     Select Case crop
        Case "Sunflower Shoots"
            ActiveCell = "SF"
        Case "Basil"
            ActiveCell = "Bas"
        Case "Pea Shoots"
            ActiveCell = "Pea"
        Case "Radish Shoots"
            ActiveCell = "Rad"
        Case "Buckwheat Shoots"
            ActiveCell = "BW"
        Case "Wheatgrass"
            ActiveCell = "WG"
    End Select
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = size
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = q
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = price
         
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = price * q
     
     ActiveCell.Offset(0, 1).Select
     ActiveCell.Font.Color = fcolor
     ActiveCell = pay
     
     ActiveCell.Offset(1, -8).Select
End Sub

Sub Fill2()
 ActiveCell.Offset(1, 0).Select
    ActiveCell = "FARM ABC"
    ActiveCell.Offset(-3, 1).Select
    ActiveCell = customer
    ActiveCell.Offset(1, 0).Select
    ActiveCell = 2
    q = q - 2
    ActiveCell.Offset(0, 1).Select
    ActiveCell = "LB"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = crop
    ActiveCell.Offset(1, 0).Select
    ActiveCell = harvestDate
    ActiveCell.Offset(1, 0).Select
    ActiveCell = "Refrigerate, Best used within the week, wash before use"

    Select Case crop
        Case "Sunflower Shoots"
            sf2pBags = sf2pBags + 1
        Case "Pea Shoots"
            pea2pBags = pea2pBags + 1
        Case "Radish Shoots"
            rad2pBags = rad2pBags + 1
        Case "Wheatgrass"
            wg2pBags = wg2pBags + 1
    End Select
End Sub

Sub Fillq()
    ActiveCell.Offset(1, 0).Select
    ActiveCell = "FARM ABC"
    ActiveCell.Offset(-3, 1).Select
    ActiveCell = customer
    ActiveCell.Offset(1, 0).Select
    ActiveCell = q
    q = 0
    ActiveCell.Offset(0, 1).Select
    ActiveCell = "LB"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = crop
    ActiveCell.Offset(1, 0).Select
    ActiveCell = harvestDate
    ActiveCell.Offset(1, 0).Select
    ActiveCell = "Refrigerate, Best used within the week, wash before use"

    Select Case crop
        Case "Sunflower Shoots"
            sf1pBags = sf1pBags + 1
        Case "Pea Shoots"
            pea1pBags = pea1pBags + 1
        Case "Radish Shoots"
            rad1pBags = rad1pBags + 1
        Case "Wheatgrass"
            wg1pBags = wg1pBags + 1
        Case "Buckwheat Shoots"
            bw1pBags = bw1pBags + 1
    End Select
End Sub

Sub Fill1()
    ActiveCell.Offset(1, 0).Select
    ActiveCell = "FARM ABC"
    ActiveCell.Offset(-3, 1).Select
    ActiveCell = customer
    ActiveCell.Offset(1, 0).Select
    ActiveCell = 1
    q = q - 1
    ActiveCell.Offset(0, 1).Select
    ActiveCell = "LB"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = crop
    ActiveCell.Offset(1, 0).Select
    ActiveCell = harvestDate
    ActiveCell.Offset(1, 0).Select
    ActiveCell = "Refrigerate, Best used within the week, wash before use"
    
    Select Case crop
        Case "Sunflower Shoots"
            sf1pBags = sf1pBags + 1
        Case "Pea Shoots"
            pea1pBags = pea1pBags + 1
        Case "Radish Shoots"
            rad1pBags = rad1pBags + 1
        Case "Wheatgrass"
            wg1pBags = wg1pBags + 1
        Case "Buckwheat shoots"
                bw1pBags = bw1pBags + 1
    End Select
    
End Sub
Sub Move2next()
    If labelColumn < 3 Then
        ActiveCell.Offset(-3, 2).Select
        labelColumn = labelColumn + 1
    Else
        ActiveCell.Offset(1, -11).Select
        labelColumn = 1
    End If
End Sub

Sub Logo()
  '  picCell = ActiveCell.Address
  '  With ActiveSheet.Pictures.Insert("C:\Users\chris\Dropbox\1 - Vancouver Food Pedalers Cooperative\Logo.jpg")
    'With ActiveSheet.Pictures.Insert("C:\Users\Daniel\Dropbox\Food Pedalers\Logo.jpg")
    'With ActiveSheet.Pictures.Insert("C:\Users\fozzarelo\Dropbox\Food Pedalers\Logo.jpg")
  '      .Left = ActiveSheet.Range(picCell).Left + 2
   '     .Top = ActiveSheet.Range(picCell).Top + 8
   '     .Width = 50
   '     .Height = 50
  '      .Placement = xlFreeFloating
   ' End With
End Sub

Sub ClearSheets()
    Sheets("LabelMaker").Select
    'Logo
    'ActiveSheet.Shapes.SelectAll
    'Selection.Delete
    ActiveSheet.Cells.ClearContents
   
    Sheets("Routes").Select
    Range("A4:I77").ClearContents
    Range("A82:I152").ClearContents
End Sub

Sub ClearRoutes2Sales()   ' not in use
    Sheets("routes2Sales").Select
    Range("A4:I65").ClearContents
    Range("A68:I125").ClearContents
    Sheets("LabelMaker").Select
End Sub

Sub Seeding()
    '
    ' Seeding Macro
    ' Finds out what to plant this week
    ' Strarting from a Tuesday "Harvest day" cell
    ' Keyboard Shortcut: Ctrl+k
    
    'Position for writing in routes J23, the go back to lbp
    Sheets("ExpectedSales").Select
    
    ActiveCell.Select
    harvestDate = ActiveCell
    ActiveCell.Offset(4, 0).Select
    ActiveCell.Offset(0, 1).Select
    
    For k = 1 To 2
        For J = 1 To 4
            For i = 1 To 9
                If ActiveCell >= harvestDate Then
                    If ActiveCell < harvestDate + 7 Then
                        'Get sow day
                        sowDay = ActiveCell
                        'Get Sow crop
                        ActiveCell.Offset(1, 0).Select
                        sowCrop = ActiveCell
                        'Get sow amount
                        ActiveCell.Offset(68, 0).Select
                        sowTrays = ActiveCell
                        ActiveCell.Offset(-69, 0).Select
                        
                        If sowTrays > 0 Then
                            'Write sow entry on "routes" J23
                            Sheets("Routes").Select
                            ActiveCell = UCase(Left(WeekdayName(Weekday(sowDay), , 1), 2))
                            ActiveCell.Offset(0, 1).Select
                            ActiveCell = sowDay
                            ActiveCell.Offset(0, 1).Select
                            ActiveCell = sowCrop
                            ActiveCell.Offset(0, 1).Select
                            ActiveCell = sowTrays
                            ActiveCell.Offset(1, -3).Select
                            Sheets("ExpectedSales").Select
                        End If
                    End If
                End If
                ActiveCell.Offset(0, 2).Select
            Next i
            ActiveCell.Offset(0, 5).Select
        Next J
        ActiveCell.Offset(0, -92).Select
        ActiveCell.Offset(87, 0).Select
    Next k
    
    ' /////Special case adds the sowing of 3 basil trays on saturday, every other week
    'oddWeek = Format(harvestDate, "ww") - 2 * Int(Format(harvestDate, "ww") / 2)
    'If oddWeek > 0 Then
        'crop = "Basil"
        'q = 2
        'AddOrder
        'Sheets("ExpectedSales").Select
        'crop = "Sunflower Shoots"
        'q = ActiveCell
        'Sheets("Routes").Select
        'ActiveCell = "SA"
        'ActiveCell.Offset(0, 1).Select
        'ActiveCell = "Saturday"
        'ActiveCell.Offset(0, 1).Select
        'ActiveCell = "Bas"
        'ActiveCell.Offset(0, 1).Select
        'ActiveCell = 3
        'ActiveCell.Offset(1, -3).Select
        'Sheets("ExpectedSales").Select
    'End If
    ' /// End of special case
    
    ActiveCell.Offset(-178, 0).Select
    
    Sheets("Routes").Select
    Range("J23:M38").Sort key1:=Range("K23:K38"), _
    order1:=xlAscending, Header:=xlNo
    Sheets("ExpectedSales").Select
    
End Sub

Sub SortRoutes()
        'Sort the routes by Route/Size/Crop
        Sheets("Routes").Select
        Range("A4:I77").Sort key1:=Range("B4:B77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A4:I77").Sort key1:=Range("E4:E77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A4:I77").Sort key1:=Range("D4:D77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("B82:B152"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("E82:E152"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("D82:D152"), _
        order1:=xlAscending, Header:=xlNo
        
        'Paint the "pay" column so that it's nearly invisible
        Range("I4:I77").Select
        With Selection.Font
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        End With
        
        Range("I82:I152").Select
        With Selection.Font
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        End With
End Sub

Sub SortRoutesbyRoute()
Attribute SortRoutesbyRoute.VB_ProcData.VB_Invoke_Func = "r\n14"
        'Sort the routes by Crop/Client/Route
        Sheets("Routes").Select
        Range("A4:I77").Sort key1:=Range("D4:D77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A4:I77").Sort key1:=Range("C4:C77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A4:I77").Sort key1:=Range("B4:B77"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("D82:D152"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("C82:C152"), _
        order1:=xlAscending, Header:=xlNo
        
        Range("A82:I152").Sort key1:=Range("B82:B152"), _
        order1:=xlAscending, Header:=xlNo
        
        'Paint the "pay" column so that it's nearly invisible
        Range("I4:I77").Select
        With Selection.Font
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        End With
        
        Range("I82:I152").Select
        With Selection.Font
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
        End With
End Sub


Sub AddBagSummay()
    'Add a summary bag count
    
    
    ActiveCell = "SF 2p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = sf2pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "SF 1p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = sf1pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "Pea 2p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = pea2pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "Pea 1p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = pea1pBags
    ActiveCell.Offset(1, -1).Select
          
    ActiveCell = "Rad 1p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = rad1pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "WG 2p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = wg2pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "WG 1p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = wg1pBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "BW 1p"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = bw1pBags
    ActiveCell.Offset(2, -1).Select
    
    ActiveCell = "SF Small"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = sfSmallBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "Pea Small"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = peaSmallBags
    ActiveCell.Offset(1, -1).Select
    
    ActiveCell = "Rad Small"
    ActiveCell.Offset(0, 1).Select
    ActiveCell = radishSmallBags
    ActiveCell.Offset(1, -1).Select
    
    Sheets("ExpectedSales").Select
End Sub




