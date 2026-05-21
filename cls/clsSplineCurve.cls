
'//
'// スプラインカーブ クラス
'//

Option Explicit

Private DataNumber_        As Double        '// データ数
Private LeastxNumber_      As Double        '// 一番小さな変数(日付)
Private LargestxNumber_    As Double        '// 一番小さな変数(日付)
Private GridxNumber_()     As Double        '// 具体的に与えた変数(日付)

'// 係数
Private CoeA_()   As Double        '// 係数a(i)
Private CoeB_()   As Double        '// 係数b(i)
Private CoeC_()   As Double        '// 係数c(i)
Private CoeD_()   As Double        '// 係数d(i)

'//
'// 3次スプラインカーブを生成する関数
'//
Public Sub genSpline(ByVal in_DataNumber As Long, _
                     ByRef in_xNumber() As Double, _
                     ByRef in_yNumber() As Double)
    
    '// インデクス
    Dim i As Long
    Dim j As Long
    
    '// データ数
    DataNumber_ = in_DataNumber
    
    '// データの設定
    ReDim GridxNumber_(1 To in_DataNumber)
    For i = 1 To in_DataNumber
        GridxNumber_(i) = in_xNumber(i)
        If i = 1 Then
            LeastxNumber_ = in_xNumber(i)
        ElseIf i = in_DataNumber Then
            LargestxNumber_ = in_xNumber(i)
        End If
    Next
    
    Dim h() As Double
    Dim v() As Double
    ReDim h(2 To in_DataNumber - 1, 2 To in_DataNumber - 1)
    ReDim v(2 To in_DataNumber - 1)
    
    '// 連立1次方程式のマトリックスを生成
    Call SetMatrix(in_xNumber, _
                   in_yNumber, _
                   h(), _
                   v())
    
    '// 連立方程式を掃き出し法で解く
    '// B()の内容が、u() に置き換わる
    
    '// 1から始まるように一般化する
    Dim MatrixA() As Double
    ReDim MatrixA(1 To in_DataNumber - 2, 1 To in_DataNumber - 2)
    Dim VectorB() As Double
    ReDim VectorB(1 To in_DataNumber - 2)
    
    For i = 1 To in_DataNumber - 2
        VectorB(i) = v(i + 1)
        For j = 1 To in_DataNumber - 2
            MatrixA(i, j) = h(i + 1, j + 1)
        Next
    Next
    
    If Gauss(in_DataNumber - 2, MatrixA(), VectorB()) = False Then
        MsgBox "行列解にエラーが出ています"
        Exit Sub
    End If
    
    '// u()を設定
    Dim u() As Double
    ReDim u(1 To in_DataNumber)
    
    For i = 1 To in_DataNumber
        If i = 1 Then
            u(i) = 0    '// 自然スプライン
        ElseIf i = in_DataNumber Then
            u(i) = 0    '// 自然スプライン
        Else
            u(i) = VectorB(i - 1)
        End If
    Next
    
    '// 係数の設定
    ReDim CoeA_(1 To DataNumber_ - 1)
    ReDim CoeB_(1 To DataNumber_ - 1)
    ReDim CoeC_(1 To DataNumber_ - 1)
    ReDim CoeD_(1 To DataNumber_ - 1)
    Dim TempX, TempY, TempU As Double
    
    For i = 1 To in_DataNumber - 1
        
        TempX = in_xNumber(i + 1) - in_xNumber(i)
        TempY = in_yNumber(i + 1) - in_yNumber(i)
        TempU = u(i + 1) - u(i)
        
        '// A
        CoeA_(i) = TempU / (6 * TempX)
        '// B
        CoeB_(i) = u(i) / 2
        '// C
        CoeC_(i) = (TempY / TempX) - TempX * (u(i + 1) + 2 * u(i)) / 6
        '// D
        CoeD_(i) = in_yNumber(i)
    Next
    
End Sub

'//
'// 連立1次方程式の行列を作成
'//
Private Sub SetMatrix(ByRef in_x() As Double, _
                      ByRef in_Y() As Double, _
                      ByRef out_H() As Double, _
                      ByRef out_U() As Double)
    
    Dim i As Long
    Dim j As Long
    Dim sh, hh As Double
    
    For i = 2 To DataNumber_ - 1
        sh = in_x(i) - in_x(i - 1)
        hh = in_x(i + 1) - in_x(i)
        
        '// B
        out_U(i) = 6 * ((in_Y(i + 1) - in_Y(i)) / hh - (in_Y(i) - in_Y(i - 1)) / sh)
        
        '// A
        For j = 2 To DataNumber_ - 1
            Select Case i
                Case j - 1
                    out_H(i, j) = hh
                Case j
                    out_H(i, j) = 2 * (sh + hh)
                Case j + 1
                    out_H(i, j) = sh
            End Select
        Next j
    Next i

End Sub




'//
'// 与えられたカーブからスプライン補間して算出する。
'//
Public Function calcSpline(ByVal in_Number As Double) As Double

    Dim GridNumber As Double
    GridNumber = getGridNumber(in_Number)
    
    If GridNumber = 0 Then
        Exit Function
    End If

    Dim GridxNumber As Double
    Dim CoeA        As Double
    Dim CoeB        As Double
    Dim CoeC        As Double
    Dim CoeD        As Double
    
    GridxNumber = GridxNumber_(GridNumber)
    CoeA = CoeA_(GridNumber)
    CoeB = CoeB_(GridNumber)
    CoeC = CoeC_(GridNumber)
    CoeD = CoeD_(GridNumber)

    calcSpline = CoeA * (in_Number - GridxNumber) ^ 3
    calcSpline = calcSpline + CoeB * (in_Number - GridxNumber) ^ 2
    calcSpline = calcSpline + CoeC * (in_Number - GridxNumber)
    calcSpline = calcSpline + CoeD

End Function


Private Function getGridNumber(ByVal in_Number As Double) As Double
    
    If LeastxNumber_ > in_Number Then
        '// 最小値未満はエラー
        getGridNumber = 0 '// "Error"
        Exit Function
        
    ElseIf LargestxNumber_ < in_Number Then
        '// 最大値超はエラー
        getGridNumber = 0 '// "Error"
        Exit Function
        
    End If
    
    Dim i As Long
    For i = 1 To DataNumber_ - 1
        If in_Number < GridxNumber_(i + 1) Then
            getGridNumber = i
            Exit Function
        End If
    Next
    
    '// 最大値超ピッタリは、最大値
    getGridNumber = DataNumber_ - 1

End Function

Private Function Gauss(ByVal in_Number As Long, _
                       ByRef in_Matrix() As Double, _
                       ByRef in_Vector() As Double) As Boolean
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim Tmp As Double
    Dim MaxRow As Long
    
On Error GoTo Err
    
    '後退代入
    For i = 1 To in_Number
        
        MaxRow = GetMaxRow(i, in_Number, in_Matrix())
        If i <> MaxRow Then
            For j = 1 To in_Number
                Tmp = in_Matrix(i, j)
                in_Matrix(i, j) = in_Matrix(MaxRow, j)
                in_Matrix(MaxRow, j) = Tmp
            Next
            Tmp = in_Vector(i)
            in_Vector(i) = in_Vector(MaxRow)
            in_Vector(MaxRow) = Tmp
        End If
        
        For j = i + 1 To in_Number
            Tmp = -1 * in_Matrix(j, i) / in_Matrix(i, i)
            in_Vector(j) = in_Vector(j) + Tmp * in_Vector(i)
            For K = i To in_Number
                in_Matrix(j, K) = in_Matrix(j, K) + Tmp * in_Matrix(i, K)
                in_Matrix(j, K) = GoZero(in_Matrix(j, K))
            Next K
        Next j
    Next i
    
    '前進消去
    in_Vector(in_Number) = in_Vector(in_Number) / in_Matrix(in_Number, in_Number)
    For i = in_Number - 1 To 1 Step -1
    
        For j = i + 1 To in_Number
            in_Vector(i) = in_Vector(i) - in_Vector(j) * in_Matrix(i, j)
        Next j
        in_Vector(i) = in_Vector(i) / in_Matrix(i, i)
    
    Next i
    
    
    Gauss = True
    Exit Function
    
Err:
    Gauss = False
End Function

Private Function GetMaxRow(ByVal in_sRow As Long, _
                           ByVal in_Number, _
                           ByRef in_Matrix() As Double) As Long

    Dim i As Long
    Dim TmpMax As Double
    TmpMax = 0
    For i = in_sRow To in_Number
        If Abs(in_Matrix(i, in_sRow)) > TmpMax Then
            TmpMax = Abs(in_Matrix(i, in_sRow))
            GetMaxRow = i
        End If
    Next
End Function

Private Function GoZero(ByRef in_Number As Double) As Double
    GoZero = in_Number
    If Abs(in_Number) < 10 ^ (-10) Then
        in_Number = 0#
    End If
End Function



