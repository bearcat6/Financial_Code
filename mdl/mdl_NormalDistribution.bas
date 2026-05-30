
Option Explicit

'// ============================================================================
'//  正規分布ユーティリティ
'// ============================================================================
'//
'//  出典:
'//      日科技連「ファイナンス工学入門 III」掲載の近似式をベースに、
'//      VBA共通部品として利用しやすいように、入力チェック・コメント・関数名を整理。
'//
'//  主な関数:
'//      StdNormPdf      : 標準正規分布の確率密度関数 φ(x)
'//      StdNormCdf      : 標準正規分布の累積分布関数 Φ(x)
'//      StdNormInv      : 標準正規分布の逆累積分布関数 Φ^-1(p)
'//      BiNormPdf       : 2変量正規分布の確率密度関数
'//      BiNormCdf       : 2変量正規分布の累積分布関数
'//
'//  互換用の旧関数名:
'//      NormSDens, NormSDist, NormSInv, NormDens2dim, NormDist2dim
'//
'//  Excel標準関数を使える場合:
'//      ワークシート上:
'//          =NORM.S.DIST(A1, TRUE)      ' 標準正規分布関数 Φ(x)
'//          =NORM.S.DIST(A1, FALSE)     ' 標準正規密度関数 φ(x)
'//          =NORM.S.INV(A1)             ' 逆標準正規分布関数 Φ^-1(p)
'//
'//      VBA上:
'//          Application.WorksheetFunction.Norm_S_Dist(x, True)
'//          Application.WorksheetFunction.Norm_S_Dist(x, False)
'//          Application.WorksheetFunction.Norm_S_Inv(p)
'//
'//      本モジュールでは、StdNormCdf / StdNormInv の第2引数に True を渡すと、
'//      Excel標準関数の利用を試み、失敗した場合は本モジュールの近似式に戻ります。
'//
'//  注意:
'//      ・本モジュールの分布関数は近似式です。
'//      ・リスク量計測、価格評価、監査対応資料等で利用する場合は、
'//        利用範囲における精度検証を行ってください。
'//      ・2変量正規密度は |rho| = 1 のとき通常の2次元密度を持たないため、
'//        BiNormPdf ではエラーとしています。
'//
'//  Last Reviewed:
'//      2026-05-30
'// ============================================================================

Private Const PI_VALUE As Double = 3.14159265358979
Private Const RHO_EPSILON As Double = 0.000000000001
'


'// ----------------------------------------------------------------------------
'//  標準正規分布の確率密度関数 φ(x)
'// ----------------------------------------------------------------------------
Public Function StdNormPdf(ByVal in_x As Double, _
                           Optional ByVal in_UseExcelFunction As Boolean = False) As Double
    If in_UseExcelFunction Then
        On Error GoTo UseApproximation
        StdNormPdf = Application.WorksheetFunction.Norm_S_Dist(in_x, False)
        Exit Function
    End If

UseApproximation:
    On Error GoTo 0
    StdNormPdf = Exp(-0.5 * in_x * in_x) / Sqr(2# * PI_VALUE)
End Function

'// ----------------------------------------------------------------------------
'//  標準正規分布の累積分布関数 Φ(x)
'// ----------------------------------------------------------------------------
Public Function StdNormCdf(ByVal in_x As Double, _
                           Optional ByVal in_UseExcelFunction As Boolean = False) As Double
    If in_UseExcelFunction Then
        On Error GoTo UseApproximation
        StdNormCdf = Application.WorksheetFunction.Norm_S_Dist(in_x, True)
        Exit Function
    End If

UseApproximation:
    On Error GoTo 0
    StdNormCdf = StdNormCdfApprox(in_x)
End Function

'// ----------------------------------------------------------------------------
'//  標準正規分布の逆累積分布関数 Φ^-1(p)
'//
'//  p は 0 < p < 1 である必要があります。
'//  p = 0 や p = 1 は理論上 ±∞ になるため、数値計算上はエラーとします。
'// ----------------------------------------------------------------------------
Public Function StdNormInv(ByVal in_p As Double, _
                           Optional ByVal in_UseExcelFunction As Boolean = False) As Double
    ValidateProbabilityOpenInterval in_p, "StdNormInv"

    If in_UseExcelFunction Then
        On Error GoTo UseApproximation
        StdNormInv = Application.WorksheetFunction.Norm_S_Inv(in_p)
        Exit Function
    End If

UseApproximation:
    On Error GoTo 0
    StdNormInv = StdNormInvApprox(in_p)
End Function

'// ----------------------------------------------------------------------------
'//  2変量正規分布の確率密度関数
'//
'//  mu1, mu2       : 平均
'//  sigma1, sigma2 : 標準偏差。0より大きい必要があります。
'//  rho            : 相関係数。-1 < rho < 1 である必要があります。
'//
'//  注意:
'//      rho = 1 または rho = -1 のとき、分布は1本の直線上に潰れるため、
'//      通常の2次元密度関数は定義されません。
'// ----------------------------------------------------------------------------
Public Function BiNormPdf(ByVal in_x1 As Double, _
                          ByVal in_x2 As Double, _
                          ByVal in_Mu1 As Double, _
                          ByVal in_Mu2 As Double, _
                          ByVal in_Sigma1 As Double, _
                          ByVal in_Sigma2 As Double, _
                          ByVal in_rho As Double) As Double

    ValidateSigma in_Sigma1, "in_sigma1", "BiNormPdf"
    ValidateSigma in_Sigma2, "in_sigma2", "BiNormPdf"
    ValidateRhoOpenInterval in_rho, "BiNormPdf"

    Dim z1 As Double
    Dim z2 As Double
    Dim oneMinusRho2 As Double

    z1 = (in_x1 - in_Mu1) / in_Sigma1
    z2 = (in_x2 - in_Mu2) / in_Sigma2
    oneMinusRho2 = 1# - in_rho * in_rho

    BiNormPdf = Exp(-((z1 * z1) - 2# * in_rho * z1 * z2 + (z2 * z2)) / (2# * oneMinusRho2)) _
                / (2# * PI_VALUE * in_Sigma1 * in_Sigma2 * Sqr(oneMinusRho2))
End Function

'// ----------------------------------------------------------------------------
'//  2変量正規分布の累積分布関数
'//
'//  戻り値:
'//      P(X1 <= x1, X2 <= x2)
'//
'//  rho = 1, rho = -1 の完全相関・完全逆相関も、分布関数としては処理します。
'// ----------------------------------------------------------------------------
Public Function BiNormCdf(ByVal in_x1 As Double, _
                          ByVal in_x2 As Double, _
                          ByVal in_Mu1 As Double, _
                          ByVal in_Mu2 As Double, _
                          ByVal in_Sigma1 As Double, _
                          ByVal in_Sigma2 As Double, _
                          ByVal in_rho As Double, _
                          Optional ByVal in_UseExcelFunctionFor1D As Boolean = False) As Double

    ValidateSigma in_Sigma1, "in_sigma1", "BiNormCdf"
    ValidateSigma in_Sigma2, "in_sigma2", "BiNormCdf"
    ValidateRhoClosedInterval in_rho, "BiNormCdf"

    Dim z1 As Double
    Dim z2 As Double

    z1 = (in_x1 - in_Mu1) / in_Sigma1
    z2 = (in_x2 - in_Mu2) / in_Sigma2

    BiNormCdf = BiNormCdfCore(z1, z2, in_rho, in_UseExcelFunctionFor1D)
End Function

'// ============================================================================
'//  互換用ラッパー
'// ============================================================================
'//  既存ブック・既存VBAで旧関数名を使用している場合に備えて残しています。
'//  新規コードでは、StdNormPdf / StdNormCdf / StdNormInv / BiNormPdf / BiNormCdf
'//  の利用を推奨します。
'// ============================================================================

Public Function NormSDens(ByVal in_Src As Double) As Double
    NormSDens = StdNormPdf(in_Src)
End Function

Public Function NormSDist(ByVal in_Src As Double) As Double
    NormSDist = StdNormCdf(in_Src)
End Function

Public Function NormSInv(ByVal in_Src As Double) As Double
    NormSInv = StdNormInv(in_Src)
End Function

Public Function NormDens2dim(ByVal in_x1 As Double, _
                             ByVal in_x2 As Double, _
                             ByVal in_Mu1 As Double, _
                             ByVal in_Mu2 As Double, _
                             ByVal in_Sigma1 As Double, _
                             ByVal in_Sigma2 As Double, _
                             ByVal in_rho As Double) As Double
    NormDens2dim = BiNormPdf(in_x1, in_x2, in_Mu1, in_Mu2, in_Sigma1, in_Sigma2, in_rho)
End Function

Public Function NormDist2dim(ByVal in_x1 As Double, _
                             ByVal in_x2 As Double, _
                             ByVal in_Mu1 As Double, _
                             ByVal in_Mu2 As Double, _
                             ByVal in_Sigma1 As Double, _
                             ByVal in_Sigma2 As Double, _
                             ByVal in_rho As Double) As Double
    NormDist2dim = BiNormCdf(in_x1, in_x2, in_Mu1, in_Mu2, in_Sigma1, in_Sigma2, in_rho)
End Function

'// ============================================================================
'//  内部処理
'// ============================================================================

Private Function StdNormCdfApprox(ByVal in_x As Double) As Double
    Const A1 As Double = 0.31938153
    Const A2 As Double = -0.356563782
    Const A3 As Double = 1.781477937
    Const A4 As Double = -1.821255978
    Const A5 As Double = 1.330274429
    Const GAMMA_VALUE As Double = 0.2316419

    Dim h As Double
    Dim absX As Double
    Dim tail As Double

    If in_x = 0# Then
        StdNormCdfApprox = 0.5
        Exit Function
    End If

    absX = Abs(in_x)
    h = 1# / (1# + GAMMA_VALUE * absX)
    tail = StdNormPdf(absX) * (A1 + (A2 + (A3 + (A4 + A5 * h) * h) * h) * h) * h

    If in_x < 0# Then
        StdNormCdfApprox = tail
    Else
        StdNormCdfApprox = 1# - tail
    End If
End Function

Private Function StdNormInvApprox(ByVal in_p As Double) As Double
    Dim y As Double
    Dim w As Double

    y = 0.5 - in_p

    If Abs(y) <= 0.42 Then
        Const A0 As Double = 2.50662823884
        Const A1 As Double = -18.61500062529
        Const A2 As Double = 41.39119773534
        Const A3 As Double = -25.44106049637

        Const B0 As Double = 1#
        Const B1 As Double = -8.4735109309
        Const B2 As Double = 23.08336743743
        Const B3 As Double = -21.06224101826
        Const B4 As Double = 3.13082909833

        w = -y * (A0 + (A1 + (A2 + A3 * y * y) * y * y) * y * y) _
             / (B0 + (B1 + (B2 + (B3 + B4 * y * y) * y * y) * y * y) * y * y)
    Else
        Dim k1 As Double
        Dim k2 As Double
        Dim z As Double
        Dim C(0 To 8) As Double
        Dim T(0 To 8) As Double
        Dim i As Long

        k1 = 0.417988642492643
        k2 = 4.24546868813766
        z = k1 * (2# * Log(-Log(0.5 - Abs(y))) - k2)

        C(0) = 7.71088707054879
        C(1) = 2.77720135336852
        C(2) = 0.3614964129261
        C(3) = 3.73418233434554E-02
        C(4) = 2.8297143036967E-03
        C(5) = 1.625716917922E-04
        C(6) = 0.000008017330474
        C(7) = 3.840919865E-07
        C(8) = 0.000000012970717

        T(0) = 1#
        T(1) = z

        For i = 2 To 8
            T(i) = 2# * z * T(i - 1) - T(i - 2)
        Next i

        w = 0.5 * C(0) * T(0)
        For i = 1 To 8
            w = w + C(i) * T(i)
        Next i

        If y > 0# Then
            w = -w
        End If
    End If

    StdNormInvApprox = w
End Function

Private Function BiNormCdfCore(ByVal in_z1 As Double, _
                               ByVal in_z2 As Double, _
                               ByVal in_rho As Double, _
                               ByVal in_UseExcelFunctionFor1D As Boolean) As Double

    Dim A(1 To 4) As Double
    Dim B(1 To 4) As Double

    A(1) = 0.325303
    A(2) = 0.4211071
    A(3) = 0.1334425
    A(4) = 0.006374323

    B(1) = 0.1337764
    B(2) = 0.6243247
    B(3) = 1.3425378
    B(4) = 2.2626645

    BiNormCdfCore = 0#

    If Abs(in_rho - 1#) <= RHO_EPSILON Then
        '// 完全相関の場合は、標準化後の小さい方で確率が決まります。
        BiNormCdfCore = StdNormCdf(MinDouble(in_z1, in_z2), in_UseExcelFunctionFor1D)

    ElseIf Abs(in_rho + 1#) <= RHO_EPSILON Then
        '// 完全逆相関の場合。数値誤差により負になることがあるため、0で下限処理します。
        BiNormCdfCore = StdNormCdf(in_z1, in_UseExcelFunctionFor1D) _
                        - StdNormCdf(-in_z2, in_UseExcelFunctionFor1D)
        BiNormCdfCore = MaxDouble(BiNormCdfCore, 0#)

    ElseIf in_z1 <= 0# And in_z2 <= 0# And in_rho <= 0# Then
        Dim t1 As Double
        Dim t2 As Double
        Dim i As Long
        Dim j As Long

        t1 = in_z1 / Sqr(2# * (1# - (in_rho * in_rho)))
        t2 = in_z2 / Sqr(2# * (1# - (in_rho * in_rho)))

        For i = 1 To 4
            For j = 1 To 4
                BiNormCdfCore = BiNormCdfCore _
                    + A(i) * A(j) _
                    * Exp(t1 * (2# * B(i) - t1) _
                        + t2 * (2# * B(j) - t2) _
                        + 2# * in_rho * (B(i) - t1) * (B(j) - t2))
            Next j
        Next i

        BiNormCdfCore = Sqr(1# - (in_rho * in_rho)) / PI_VALUE * BiNormCdfCore

    ElseIf in_z1 * in_z2 * in_rho <= 0# Then
        If in_z1 <= 0# And in_z2 >= 0# And in_rho >= 0# Then
            BiNormCdfCore = StdNormCdf(in_z1, in_UseExcelFunctionFor1D) _
                            - BiNormCdfCore(in_z1, -in_z2, -in_rho, in_UseExcelFunctionFor1D)

        ElseIf in_z1 >= 0# And in_z2 <= 0# And in_rho >= 0# Then
            BiNormCdfCore = StdNormCdf(in_z2, in_UseExcelFunctionFor1D) _
                            - BiNormCdfCore(-in_z1, in_z2, -in_rho, in_UseExcelFunctionFor1D)

        ElseIf in_z1 >= 0# And in_z2 >= 0# And in_rho <= 0# Then
            BiNormCdfCore = StdNormCdf(in_z1, in_UseExcelFunctionFor1D) _
                            + StdNormCdf(in_z2, in_UseExcelFunctionFor1D) _
                            - 1# _
                            + BiNormCdfCore(-in_z1, -in_z2, in_rho, in_UseExcelFunctionFor1D)
        End If

    Else
        Dim denominator As Double
        Dim rho1 As Double
        Dim rho2 As Double
        Dim delta As Double

        denominator = Sqr((in_z1 * in_z1) - 2# * in_rho * in_z1 * in_z2 + (in_z2 * in_z2))

        '// 通常はここで denominator = 0 になりませんが、丸め誤差対策として保険を入れます。
        If denominator = 0# Then
            BiNormCdfCore = BiNormCdfCore(in_z1, in_z2, 0#, in_UseExcelFunctionFor1D)
            Exit Function
        End If

        rho1 = (in_rho * in_z1 - in_z2) * Sgn(in_z1) / denominator
        rho2 = (in_rho * in_z2 - in_z1) * Sgn(in_z2) / denominator
        delta = (1# - Sgn(in_z1) * Sgn(in_z2)) / 4#

        BiNormCdfCore = BiNormCdfCore(in_z1, 0#, rho1, in_UseExcelFunctionFor1D) _
                        + BiNormCdfCore(0#, in_z2, rho2, in_UseExcelFunctionFor1D) _
                        - delta
    End If

    '// 近似計算・丸め誤差により、ごく稀に [0,1] をわずかに外れる場合の補正。
    BiNormCdfCore = ClampProbability(BiNormCdfCore)
End Function

Private Sub ValidateProbabilityOpenInterval(ByVal in_p As Double, ByVal in_functionName As String)
    If in_p <= 0# Or in_p >= 1# Then
        Err.Raise 5, in_functionName, in_functionName & ": 確率 p は 0 より大きく 1 より小さい値を指定してください。"
    End If
End Sub

Private Sub ValidateSigma(ByVal in_sigma As Double, ByVal in_argumentName As String, ByVal in_functionName As String)
    If in_sigma <= 0# Then
        Err.Raise 5, in_functionName, in_functionName & ": " & in_argumentName & " は 0 より大きい標準偏差を指定してください。"
    End If
End Sub

Private Sub ValidateRhoClosedInterval(ByVal in_rho As Double, ByVal in_functionName As String)
    If in_rho < -1# Or in_rho > 1# Then
        Err.Raise 5, in_functionName, in_functionName & ": 相関係数 rho は -1 以上 1 以下を指定してください。"
    End If
End Sub

Private Sub ValidateRhoOpenInterval(ByVal in_rho As Double, ByVal in_functionName As String)
    If in_rho <= -1# Or in_rho >= 1# Then
        Err.Raise 5, in_functionName, in_functionName & ": 密度関数では相関係数 rho は -1 より大きく 1 より小さい値を指定してください。"
    End If
End Sub


Private Function MaxDouble(ByVal in_x As Double, ByVal in_y As Double) As Double
    If in_x < in_y Then
        MaxDouble = in_y
    Else
        MaxDouble = in_x
    End If
End Function

Private Function MinDouble(ByVal in_x As Double, ByVal in_y As Double) As Double
    If in_x < in_y Then
        MinDouble = in_x
    Else
        MinDouble = in_y
    End If
End Function

Private Function ClampProbability(ByVal in_p As Double) As Double
    If in_p < 0# Then
        ClampProbability = 0#
    ElseIf in_p > 1# Then
        ClampProbability = 1#
    Else
        ClampProbability = in_p
    End If
End Function

'// ----------------------------------------------------------------------------
'//  簡易テスト用
'// ----------------------------------------------------------------------------
'//  VBEのイミディエイトウィンドウで TestNormalDistribution を実行すると、
'//  代表的な値を確認できます。
'// ----------------------------------------------------------------------------
Public Sub TestNormalDistribution()
    Debug.Print "StdNormPdf(0)       = "; StdNormPdf(0#); "  expected about 0.3989422804"
    Debug.Print "StdNormCdf(0)       = "; StdNormCdf(0#); "  expected 0.5"
    Debug.Print "StdNormInv(0.5)     = "; StdNormInv(0.5); "  expected 0"
    Debug.Print "StdNormInv(0.99)    = "; StdNormInv(0.99); "  expected about 2.326347874"
    Debug.Print "BiNormCdf(0,0,rho=0)= "; BiNormCdf(0#, 0#, 0#, 0#, 1#, 1#, 0#); "  expected 0.25"
    Debug.Print "BiNormCdf(0,0,rho=.5)= "; BiNormCdf(0#, 0#, 0#, 0#, 1#, 1#, 0.5); "  expected about 0.3333333333"
End Sub


