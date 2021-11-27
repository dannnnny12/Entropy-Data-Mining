VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12855
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Search_B 
      Caption         =   "Search Backward"
      Height          =   255
      Left            =   10200
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Search_F 
      Caption         =   "Search Forward"
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox entro_list 
      Height          =   5280
      Left            =   4920
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ListBox Bwd 
      Height          =   5100
      Left            =   9600
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ListBox Fwd 
      Height          =   5100
      Left            =   7440
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ListBox ef_list 
      Height          =   5280
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ListBox ew_list 
      Height          =   5280
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Entropy"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Search Forward/Backward"
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Equal Frequency"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Equal Width"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fullData(106, 10), numerator, denominator As Double
Dim copyData(106, 10) As Double
Dim sortedCol(106) As Double
Dim tempU, sort_cla(106) As Double
Dim tempCount(10) As Double
Dim attributeInd(9), index As Integer
Dim col(10) As Double
Dim chosen(9), attributeValue As Integer
Dim tempSplitPoint(9), EF_split(9) As Double '暫存分割點
Dim SortClassValue(106, 10), SortAttribute(106, 10), ord_class(106) As Double
Dim EWData(106, 10) As Double
Dim EFData(106, 10) As Double
Dim EBData(106, 10) As Double
Dim splitPointArray(100) As Double
Dim arrdata(106, 10) As Double
Dim EWUArray(10, 10) As Double
Dim EFUArray(10, 10) As Double
Dim EBUArray(10, 10) As Double


'H(x)
Private Function H(arrdata, col) As Double
    num_count = 0
    Dim Hx As Double
    '算categorical各值的數量
    For attributeValue = 0 To 9
        For index = 1 To 106
            If arrdata(index, col) = attributeValue Then
                num_count = num_count + 1
            End If
        Next index
        tempCount(attributeValue) = num_count
        num_count = 0
    Next attributeValue
    Hx = 0
    For j = 0 To 9
        px = tempCount(j) / 106
        Hx = Hx + -px * Log2(px)
    Next j
    'HArray(col) = Hx
    H = Hx
End Function

'H(X,Y)
Public Function Hxy(arrdata, col1, col2)
   num_count = 0
    For x = 0 To 9
        For y = 0 To 9
            For index = 1 To 106
                If arrdata(index, col1) = x And arrdata(index, col2) = y Then
                    num_count = num_count + 1
                End If
            Next index
            Pxy = num_count / 106
            tempHxy = tempHxy + -Pxy * Log2(Pxy)
            num_count = 0
        Next y
    Next x
    Hxy = tempHxy
End Function

'U
Private Function U(tempU, data, col1, col2)
        If (H(data, col1) + H(data, col2)) = 0 Then
        tempU(col1, col2) = 1
        Else
        u_result = 2 * ((H(data, col1) + H(data, col2) - Hxy(data, col1, col2)) / (H(data, col1) + H(data, col2)))
        tempU(col1, col2) = u_result
        End If
End Function

Function Goodness(arrdata)
    numerator = 0
    denominator = 0
    'numerator
    For i = 1 To 9
        If chosen(i) = 1 Then
       numerator = numerator + arrdata(i, 10)

        End If
    Next i
    'denominator
    For i = 1 To 9
        For j = 1 To 9
            If chosen(i) = 1 And chosen(j) = 1 Then
            denominator = denominator + arrdata(i, j)
            End If
        Next j
    Next i
    
    
    denominator = Sqr(denominator)
    If denominator <> 0 Then
        Goodness = numerator / (denominator)
        
    Else
    Goodness = 0 '分母為0 goodness=0
    End If
End Function

'Sorting
Sub Sorting(col)
Dim i, j As Integer
Dim temp As Double

For k = 1 To 106
    sortedCol(k) = fullData(k - 1, col - 1)
Next k

'bubble sort
arrMin = LBound(sortedCol)
arrMax = UBound(sortedCol)
For i = arrMin To arrMax - 1
       For j = i + 1 To arrMax
           If sortedCol(i) > sortedCol(j) Then
               temp = sortedCol(i)
               sortedCol(i) = sortedCol(j)
               sortedCol(j) = temp
           End If
        Next j
    Next i
End Sub

Sub equal_width(col1 As Integer)
Dim w As Double

Dim split_point(9) As Double

    w = (sortedCol(106) - sortedCol(1)) / 10  'ten bins
    ew_list.AddItem ("Column" & "  " & col1)
    ew_list.AddItem "W= " & w

For j = 0 To 9

    upperBound = sortedCol(1) + w * j
    lowerBound = sortedCol(1) + w * (j + 1)
    tempSplitPoint(j) = lowerBound
    ew_list.AddItem "interval" & (j + 1) & ":"
    ew_list.AddItem (upperBound & " " & "," & " " & lowerBound)
   
Next j

     'discrete
     For j = 1 To 106
        For k = 9 To 1 Step -1
            If EWData(j, col1) > tempSplitPoint(k - 1) And EWData(j, col1) <= tempSplitPoint(k) Then
                EWData(j, col1) = k
                Exit For
            ElseIf EWData(j, col1) <= tempSplitPoint(0) Then
                EWData(j, col1) = 0
            End If
            'test.AddItem EWData(j, col1)
        Next k
    Next j
     
    
     

End Sub

Sub equal_frequency(col As Integer)


Dim tencount As Integer
Dim elecount As Integer

    
    ef_list.AddItem ("Column" & col)
    For m = 10 To 40 '第0~40筆中找split point 以10筆資料為一區間 10*4
        tencount = m / 9 'tencount=1,2,3,4
        EF_split(tencount - 1) = ((sortedCol(m) + sortedCol(m + 1)) / 2)
        m = m + 9
    Next m
        
    
    For n = 41 To 96 '第41~96筆中找split point 以11筆資料為一區間 11*6
    elecount = n \ 10 '無條件捨去 elecount=5,6,7,8,9,10
        EF_split(elecount - 1) = ((sortedCol(n - 1) + sortedCol(n)) / 2)
        n = n + 10
    Next n
    
    'show interval
    For i = 0 To 9
        ef_list.AddItem "interval" & " " & (i + 1) & ":"
        If i = 0 Then
             lowerBound = sortedCol(1)
             upperBound = EF_split(i)
        ElseIf i = 9 Then
            lowerBound = EF_split(i - 1)
            upperBound = sortedCol(106)
        Else
            upperBound = EF_split(i)
            lowerBound = EF_split(i - 1)
        End If
         
        ef_list.AddItem (lowerBound & " , " & upperBound)
    Next i
    
     For j = 1 To 106
        For k = 8 To 1 Step -1
            If EFData(j, col) > EF_split(k - 1) And EFData(j, col) <= EF_split(k) Then
                EFData(j, col) = k
                Exit For
            ElseIf EFData(j, col) <= EF_split(0) Then
                EFData(j, col) = 0
                Exit For
            ElseIf EFData(j, col) > EF_split(8) Then
                EFData(j, col) = 9
                Exit For
            End If
        Next k
       
    Next j
    
    ef_list.AddItem "-------------------"

End Sub


Private Sub SortClass(col1 As Integer)

    For i = 1 To 106
        sort_cla(i) = fullData(i - 1, col1 - 1)
        ord_class(i) = fullData(i - 1, 9)
        
    Next i

    For i = 1 To 106
        For j = 1 To 105
            If sort_cla(j) > sort_cla(j + 1) Then
                temp1 = sort_cla(j)
                sort_cla(j) = sort_cla(j + 1)
                sort_cla(j + 1) = temp1
                
                '排序class值
                temp2 = ord_class(j)
                ord_class(j) = ord_class(j + 1)
                ord_class(j + 1) = temp2
            End If
        Next j
    Next i
    
    For i = 1 To 106
        SortClassValue(i, col1) = ord_class(i)
        SortAttribute(i, col1) = sort_cla(i)
    Next i
    
    entropy col1, ord_class '呼叫entropy &print
    
End Sub

Sub entropy(col1 As Integer, tempArray() As Double)
    Dim MinEnt, cutNum, TEnt, tempArraySize, stop0, stop1, stop2 As Integer
    Dim left() As Double
    Dim right() As Double
  Dim gainValue, threshold As Double
     
   
    'entro_list.AddItem "2123"
    
    MinEnt = 99999999 '紀錄最小的entropy
   
    TEnt = 0
    tempArraySize = UBound(tempArray) - LBound(tempArray)
    cutNum = 0
    
    '計算切幾個點
    If tempArraySize > 1 Then
        For i = 1 To tempArraySize - 1
            If tempArray(i) <> tempArray(i + 1) Then
                cutNum = cutNum + 1
            End If
        Next i
    End If
   
    If cutNum = 0 Then
    Else
        For i = 1 To tempArraySize
            If tempArray(i) = 1 Then
                Tcount1 = Tcount1 + 1
            ElseIf tempArray(i) = 2 Then
                Tcount2 = Tcount2 + 1
            ElseIf tempArray(i) = 3 Then
                Tcount3 = Tcount3 + 1
            ElseIf tempArray(i) = 4 Then
                Tcount4 = Tcount4 + 1
            ElseIf tempArray(i) = 5 Then
                Tcount5 = Tcount5 + 1
            ElseIf tempArray(i) = 6 Then
                Tcount6 = Tcount6 + 1
            End If
        Next i
     
        '計算Ent(S) '總set
        Tsum = Tcount1 + Tcount2 + Tcount3 + Tcount4 + Tcount5 + Tcount6
        prob = Tcount1 / (Tsum)
        TEnt = TEnt - prob * Log2(prob) 'prob被複寫，從1-6
        prob = Tcount2 / (Tsum)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tcount3 / (Tsum)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tcount4 / (Tsum)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tcount5 / (Tsum)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tcount6 / (Tsum)
        TEnt = TEnt - prob * Log2(prob)
        
        '計算左右個數
        For j = 1 To tempArraySize - 1
            If tempArray(j) <> tempArray(j + 1) Then
                splitPoint = (SortAttribute(j, col1) + SortAttribute(j + 1, col1)) / 2
            L1 = 0
            L2 = 0
            L3 = 0
            L4 = 0
            L5 = 0
            L6 = 0
                
            R1 = 0
            R2 = 0
            R3 = 0
            R4 = 0
            R5 = 0
            R6 = 0
                
                For k = 1 To tempArraySize
                    If k < j + 1 Then
                       If tempArray(k) = 1 Then
                            L1 = L1 + 1
                        ElseIf tempArray(k) = 2 Then
                            L2 = L2 + 1
                        ElseIf tempArray(k) = 3 Then
                            L3 = L3 + 1
                        ElseIf tempArray(k) = 4 Then
                            L4 = L4 + 1
                        ElseIf tempArray(k) = 5 Then
                            L5 = L5 + 1
                        ElseIf tempArray(k) = 6 Then
                            L6 = L6 + 1
                        End If
                    Else
                       If tempArray(k) = 1 Then
                            R1 = R1 + 1
                        ElseIf tempArray(k) = 2 Then
                            R2 = R2 + 1
                        ElseIf tempArray(k) = 3 Then
                            R3 = R3 + 1
                        ElseIf tempArray(k) = 4 Then
                            R4 = R4 + 1
                        ElseIf tempArray(k) = 5 Then
                            R5 = R5 + 1
                        ElseIf tempArray(k) = 6 Then
                            R6 = R6 + 1
                            End If
                    End If
                Next k
                
                LEnt = 0
                REnt = 0
              'left side s calculation
                Lsum = L1 + L2 + L3 + L4 + L5 + L6
                prob = L1 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                prob = L2 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                prob = L3 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                prob = L4 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                prob = L5 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                prob = L6 / Lsum
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                'right side s calculation
                Rsum = R1 + R2 + R3 + R4 + R5 + R6
                prob = R1 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                prob = R2 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                prob = R3 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                prob = R4 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                prob = R5 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                prob = R6 / Rsum
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                Ent = LEnt * (Lsum) / (Lsum + Rsum) + REnt * (Rsum) / (Lsum + Rsum)
            
                '停止條件
                stop0 = 6
                stop1 = 6
                stop2 = 6
                delta = Log2(3 ^ stop0 - 2) - stop0 * TEnt + stop1 * LEnt + stop2 * REnt
                reject = TEnt - Ent - (Log2(tempArraySize - 1) + delta) / tempArraySize
'                List6.AddItem "reject" & reject

                If MinEnt >= Ent And reject > 0 Then
                    MinEnt = Ent
                    MinSplitPoint = splitPoint
                    splitIndex = Lsum + 1
                End If
            End If
        Next j
        
        If MinEnt <> 99999999 Then
            splitLen = splitLen + 1
            splitPointArray(splitLen) = MinSplitPoint
            
            index_left = splitIndex - 1
            index_right = tempArraySize - index_left
            
            ReDim left(index_left) As Double
            ReDim right(index_right) As Double
            
            l = 1
            r = 1
            For i = 1 To tempArraySize
                If i < splitIndex Then
                    left(l) = tempArray(i)
                    l = l + 1
                Else
                    right(r) = tempArray(i)
                    r = r + 1
                End If
            Next i
            
            'print Entropy interval
            entro_list.AddItem "Column" & col1
            entro_list.AddItem "SplitPoint " & splitPointArray(splitLen)
            entro_list.AddItem "-----------------"
            

            For i = 1 To 106
                If EBData(i, col1) <= splitPointArray(splitLen) Then
                    EBData(i, col1) = 0
                ElseIf EBData(i, col1) > splitPointArray(splitLen) Then
                    EBData(i, col1) = 1
                End If
            Next i

            entropy col1, left
            entropy col1, right
            
        End If
    End If
End Sub


'把 log轉為2為底
Static Function Log2(x) As Double
    If x <> 0 Then
        Log2 = Log(x) / Log(2#)
    ElseIf x = 0 Then
        Log2 = 0
    End If
End Function
' Forward Function
Private Sub Forward(UArr)
    maxGoodness = 0
        
        For i = 0 To 9
            chosen(i) = 0 '初始化，全部設為0
        Next i
    'forward
        For j = 1 To 9
                index = -1 '當index=-1,停止
            For i = 1 To 9
                If chosen(i) <> 1 Then
                    chosen(i) = 1
                    tempGoodness = Goodness(UArr)
                    
                    If tempGoodness > maxGoodness Then
                        maxGoodness = tempGoodness
                        index = i
                    End If
                    chosen(i) = 0
                    tempGoodness = 0
                    End If
            Next i
                If index = -1 Then Exit For
                chosen(index) = 1
        Next j
        
                     Fwd.AddItem ("Attribute Subsets:")
                For b = 1 To 9
                    If chosen(b) = 1 Then
                        Fwd.AddItem ("A" & b)
                    End If
                Next b
                    
                  Fwd.AddItem ("Goodness：" & Math.Round(maxGoodness, 5))
 
End Sub

Private Sub Backward(tempU)
    
 For i = 0 To 9
        chosen(i) = 1
    Next i
    
    maxGoodness = 0 '初始化最大的goodness
    For n = 1 To 9
        index = -1
        For m = 1 To 9
            If chosen(m) = 1 Then
                chosen(m) = 0
                tempGoodness = Goodness(tempU)
                If tempGoodness > maxGoodness Then
                    maxGoodness = tempGoodness '存最大的goodness
                    index = m
                End If
                chosen(m) = 1
            End If
        Next m
        If index = -1 Then Exit For 'goodness下降則停止
        chosen(index) = 0
        r = index
        Bwd.AddItem ("Attribute removed：A" & r)
        Bwd.AddItem ("Goodness：" & Math.Round(maxGoodness, 5))
    Next n
    Bwd.AddItem ("     ")
    For i = 1 To 9
        If chosen(i) = 1 Then
            Bwd.AddItem ("Attribute subset A" & i)
        End If
    Next i
    
  
End Sub



Private Sub Form_Load() '讀入時已計算equal frequecy, equal width, entropy
    Dim i As Integer
    Dim path As String
    i = 0
    Open App.path & "\Breast.txt" For Input As #1
    
    
    Do While Not EOF(1)
        Line Input #1, rawTmp
        For j = 0 To 9
           Dim s As Variant
            s = Split(rawTmp, ",")
                   If j = 9 Then
                    If s(j) = "car" Then
                        fullData(i, j) = 1
                    ElseIf s(j) = "fad" Then
                        fullData(i, j) = 2
                    ElseIf s(j) = "mas" Then
                        fullData(i, j) = 3
                    ElseIf s(j) = "gla" Then
                        fullData(i, j) = 4
                    ElseIf s(j) = "con" Then
                        fullData(i, j) = 5
                    ElseIf s(j) = "adi" Then
                        fullData(i, j) = 6
                    End If
                    Else
                        fullData(i, j) = Val(s(j))
                   End If
                
            'test.AddItem (fullData(i, j))
        Next
        i = i + 1
    Loop
    Close #1
    
    
    
     For i = 1 To 106
        For j = 1 To 10
           
            EWData(i, j) = fullData(i - 1, j - 1) '需減1因為fullData是從(1,1)開始記錄
            EFData(i, j) = fullData(i - 1, j - 1)
            EBData(i, j) = fullData(i - 1, j - 1)
        Next j
    Next i
    'test.AddItem (EBData(1, 2))
    For i = 1 To 9
        Sorting (i)
        equal_width (i)
        ew_list.AddItem "----------------------"
    Next i
   
  
    For j = 1 To 9
            Sorting (j)
            equal_frequency (j)
    Next j
    
    For j = 1 To 9
            SortClass (j)
        Next j
    
   
    
    
    'equal_width
    'For n = 0 To 105
        'For m = 0 To 9
            'test.AddItem EWData(n, m) 'EWData EFData OKAY
        'Next m
    'Next n
  For m = 1 To 10
    For n = 1 To 10
        U EWUArray, EWData, m, n
        'test.AddItem EWUArray(m, n)
        U EFUArray, EFData, m, n
        'test.AddItem EFUArray(m, n)
        U EBUArray, EBData, m, n
    Next n
  Next m
  
  
End Sub


Private Sub Search_B_Click()

    Bwd.AddItem "--------Equal-width--------"
    Backward (EWUArray) '顯示在forward list上
    Fwd.AddItem ""
    
    Bwd.AddItem "--------Equal-frequency--------"
    Backward (EFUArray) '顯示在forward list上
    Fwd.AddItem ""
    
    Bwd.AddItem "--------Entropy--------"
    Backward (EBUArray) '顯示在forward list上
End Sub

Private Sub Search_F_Click()
      
    
    Fwd.AddItem "--------Equal-width--------"
    Forward (EWUArray) '顯示在backward list上
    Fwd.AddItem " "
  
    Fwd.AddItem "--------Equal-frequency--------"
    Forward (EFUArray) '顯示在backward list上
    Fwd.AddItem " "
    
    Fwd.AddItem "--------Entropy--------"
    Forward (EBUArray) '顯示在backward list上
    
End Sub
