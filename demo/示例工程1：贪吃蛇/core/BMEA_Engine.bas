Attribute VB_Name = "BMEA_Engine"
'   Emerald 相关代码
'================================================================================
'   黑嘴加密算法
'   制作: Error404
'   注意事项：
'       1.本加密算法是不可逆算法
'================================================================================
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'================================================================================
'   主体
'   <Inputs:需要加密的内容>
    Function BMEA(ByVal Inputs As String, Optional BMKey) As String
        Dim StrEA() As Byte, Key As String, temp As Long, KeyP As Integer
        Dim LongEA As Long, LongRet As String
        Dim RepLst1 As String, RepLst2() As String, RepLst As String, RepP As Integer, WaitChr As String, RepRet As String
        Dim TryRep As Long, SinStep As Long
        Dim BowRep As String, WowRep As String, BowRet As String, WowRet As String
        Dim I As Integer, S As Integer
        
        BowRep = "ABCDEF"
        WowRep = "汪嗷呜吼嘤喵"
        If IsMissing(BMKey) = False Then
            If BMKey <> "" Then WowRep = BMKey
        End If
        StrEA = Inputs & WowRep: Key = Len(Inputs): KeyP = 1
        RepLst1 = "0123456789ABCDEF"
        ReDim RepLst2(Len(RepLst1))
        RepP = 1
        For I = 1 To Len(RepLst1)
            WaitChr = Mid(RepLst1, I, 1)
            Do While RepLst2(RepP) <> ""
                SinStep = Int(Abs(Sin(I * Len(Inputs))))
                If SinStep = 0 Then SinStep = 1
                RepP = RepP + IIf(TryRep <= 404, SinStep, 1)
                If RepP > Len(RepLst1) Then RepP = RepP Mod Len(RepLst1): TryRep = TryRep + 1
                If RepP = 0 Then RepP = 1
            Loop
            RepLst2(RepP) = WaitChr
        Next
        
        For I = 1 To UBound(RepLst2)
            RepLst = RepLst & RepLst2(I)
        Next
        
        For I = 0 To UBound(StrEA)
            temp = StrEA(I) + ((Val(Mid(Key, KeyP, 1)) * Val(Mid(Key, KeyP, IIf(Len(Key) > 1, 2, 1)))) Mod 233)
            temp = temp + Len(WowRep) * Sin(UBound(StrEA) - I) - Abs(Asc(Mid(WowRep, I Mod Len(WowRep) + 1, 1))) * Cos(I)
            temp = Abs(temp) Mod 255
            StrEA(I) = temp
            KeyP = KeyP + 1
            If KeyP > Len(Key) - 1 Then KeyP = 1
        Next
        
        If (UBound(StrEA) + 1) Mod 4 <> 0 Then
            ReDim Preserve StrEA(Int(UBound(StrEA) / 4) * 4 + 4 - 1)
        End If
        
        For I = 0 To UBound(StrEA) Step 4
            CopyMemory LongEA, StrEA(I), 4
            LongRet = LongRet & Hex(LongEA)
        Next
        
        StrEA = LongRet
        
        For I = 1 To Len(LongRet)
            WaitChr = Mid(LongRet, I, 1)
            For S = 1 To Len(RepLst)
                If Mid(RepLst1, S, 1) = WaitChr Then RepRet = RepRet & Mid(RepLst, S, 1): Exit For
            Next
        Next
        
        Dim Buff As String
        For I = 1 To Len(RepRet)
            WaitChr = Mid(RepRet, I, 1)
            If (Asc(WaitChr) >= 48 And Asc(WaitChr) <= 57) And Val(Buff) <= 32767 Then
                Buff = Buff & WaitChr
            Else
                If Buff <> "" Then WowRet = WowRet & Hex(Val(Buff) Mod 32767)
                If Val(Buff) > 32767 Then WowRet = WowRet & Hex(Val(Buff) - 32767)
                WowRet = WowRet & WaitChr
                Buff = ""
            End If
        Next

        BMEA = WowRet
    End Function
    Public Function GetBMKey() As String
        Randomize
        GetBMKey = Hex(Int(Rnd * 1000000000 + 10000000)) & Hex(Int(Rnd * 1000000000 + 10000000)) & Hex(Int(Rnd * 1000000000 + 10000000)) & Hex(Int(Rnd * 1000000000 + 10000000))
    End Function
'================================================================================
