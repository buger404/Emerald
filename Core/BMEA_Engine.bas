Attribute VB_Name = "BMEA_Engine"
'   Emerald 相关代码
'================================================================================
'   黑嘴加密算法
'   制作: Error404
'   版本: 1.0 / 19.x
'   注意事项：
'       1.本加密算法是不可逆算法
'================================================================================
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'================================================================================
'   主体
'   <Inputs:需要加密的内容>
    Function BMEA(ByVal Inputs As String) As String
        Dim StrEA() As Byte, Key As String, temp As Long, KeyP As Integer
        Dim LongEA As Long, LongRet As String
        Dim RepLst1 As String, RepLst2() As String, RepLst As String, RepP As Integer, WaitChr As String, RepRet As String
        Dim TryRep As Long, SinStep As Long
        Dim BowRep As String, WowRep As String, BowRet As String, WowRet As String, Bowed As Boolean
        
        BowRep = "ABCDEF"
        WowRep = "汪嗷呜吼嘤喵"
        StrEA = Inputs: Key = Len(Inputs): KeyP = 1
        RepLst1 = "0123456789ABCDEF"
        ReDim RepLst2(Len(RepLst1))
        RepP = 1
        For i = 1 To Len(RepLst1)
            WaitChr = Mid(RepLst1, i, 1)
            Do While RepLst2(RepP) <> ""
                SinStep = Int(Abs(Sin(i * Len(Inputs))))
                If SinStep = 0 Then SinStep = 1
                RepP = RepP + IIf(TryRep <= 404, SinStep, 1)
                If RepP > Len(RepLst1) Then RepP = RepP Mod Len(RepLst1): TryRep = TryRep + 1
                If RepP = 0 Then RepP = 1
            Loop
            RepLst2(RepP) = WaitChr
        Next
        
        For i = 1 To UBound(RepLst2)
            RepLst = RepLst & RepLst2(i)
        Next
        
        For i = 0 To UBound(StrEA)
            temp = StrEA(i) + ((Val(Mid(Key, KeyP, 1)) * Val(Mid(Key, KeyP, IIf(Len(Key) > 1, 2, 1)))) Mod 233)
            temp = temp Mod 255
            StrEA(i) = temp
            KeyP = KeyP + 1
            If KeyP > Len(Key) - 1 Then KeyP = 1
        Next
        
        If (UBound(StrEA) + 1) Mod 4 <> 0 Then
            ReDim Preserve StrEA(Int(UBound(StrEA) / 4) * 4 + 4 - 1)
        End If
        
        For i = 0 To UBound(StrEA) Step 4
            CopyMemory LongEA, StrEA(i), 4
            LongRet = LongRet & Hex(LongEA)
        Next
        
        StrEA = LongRet
        
        For i = 1 To Len(LongRet)
            WaitChr = Mid(LongRet, i, 1)
            For s = 1 To Len(RepLst)
                If Mid(RepLst1, s, 1) = WaitChr Then RepRet = RepRet & Mid(RepLst, s, 1): Exit For
            Next
        Next
        
        WowRet = "..."
        TryRep = 0
        
        Do While WowRet <> ""
            WowRet = ""
        
            For i = 1 To Len(RepRet)
                WaitChr = Mid(RepRet, i, 1)
                Bowed = False
                For s = 1 To Len(BowRep)
                    If Mid(BowRep, s, 1) = WaitChr Then
                        BowRet = BowRet & Mid(WowRep, s, 1)
                        Bowed = True
                        Exit For
                    End If
                Next
                If Bowed = False Then
                    WowRet = WowRet & WaitChr
                End If
            Next
            
            RepRet = ""
            For i = 1 To Len(WowRet) Step 4
                RepRet = RepRet & Hex(Val(Mid(WowRet, i, 4)))
            Next
            TryRep = TryRep + 1
            
            If TryRep >= 233 Then Exit Do
        Loop
        
        BMEA = BowRet & WowRet
    End Function
'================================================================================
