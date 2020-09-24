Attribute VB_Name = "modRiff"
' This module provides the necessary tools for generating
'   basic RIFF (WAV) files in a standard PCM format from
'   multiple sine waves

' Code module and project created by "Urthman"
'   http://www.jsent.biz/urthman/
'   http://www.mp3.com/urthman/

Option Explicit

'   Bandwidth Selection ENUM

Enum BandWidth
    NotSet = &H0            ' Not Set
    EightBit = &H8          ' 8-bit
    SixteenBit = &H10       ' 16-bit
    TwentyFourBit = &H18    ' 24-bit
End Enum

'   Internal data structures for sample streams

Private Type WaveForm1
    Index As Long
    Count As Long
    Sample() As Double
End Type

Private Type WaveForm2
    Index As Long
    Count As Long
    Sample() As Long
End Type

'=====================================================================================
'   RIFF file format data chunks
'=====================================================================================

'   RIFF chunk
    
Private Type RiffHead
    Name As String * 4
    Size As Long
    Wave As String * 4
End Type

'   Format Chunk

Private Type FormatChunk
    Name As String * 4
    Size As Long
    AudioFormat As Integer
    Channels As Integer
    SampleRate As Long
    ByteRate As Long
    BlockAlign As Integer
    BitsPerSample As Integer
End Type

'   Wave Data Chunk

Private Type DataChunk
    Name As String * 4
    Size As Long
    Data() As Byte
End Type

'=====================================================================================

'   Temporary File variables

Dim PrtSize() As Long
Dim PrtName() As String
Dim PrtNumb As Long
Dim PrtIndx As Integer

'   The RIFF structure data

Dim RIF As RiffHead
Dim FMT As FormatChunk
Dim WAV As DataChunk

'   Variables established by InitRiff - do not alter these values
'   from the outside ... please read-only

Global Vmax As Double       ' Peak Value or preferred null value for silence
Global bWidth As BandWidth  ' Resolution - bits per sample (Long)
Global bRate As Long        ' Sample Rate

Dim WavIndx As Integer

'   Working Waveform Data

Dim Wave() As WaveForm1
Dim WrkWave As WaveForm1

'   RIF File Name - for reference

Global RifName As String

'   Mixed output waveform

Global OutWave As WaveForm2
' Attenuate: This routine applies an attenuation value against a
'   given wave generated by MakeSine. See also AttenuationValue
'
Sub Attenuate(Which%, ByVal Level As Double)

'   Which = which sine wave array
'   Level = (from AttenuationValue) the new amplitude bandwidth

Dim AdjustBy As Double
Dim aIndx&

    AdjustBy = Level / ((2 ^ bWidth) - 1)
    
    For aIndx = 0 To UBound(Wave(Which).Sample)
        Wave(Which).Sample(aIndx) = (Wave(Which).Sample(aIndx) * AdjustBy)
    Next

End Sub
' AttenuationValue: Given the decibel value to reduce a given signal by
'   this function produces the necessary sample band-width peak value
'
Function AttenuationValue(ByVal Decibel As Double) As Double

'   Decibel = the decibel value relative to the full bandwidth
'   Returns the equivalent attenuation band width

'   AttenuationValue(0) = no AttenuationValue = maximum volume
'   AttenuationValue(2.93) = reduction value for -2.93 decibels

Dim DB As Double

' To simplify programming, we'll force a negative value for adjustment
    
    DB = (Abs(Decibel) * -1)
    AttenuationValue = (10 ^ (DB / 20)) * ((2 ^ bWidth) - 1)

End Function
' ClearWaves Initialize the arrays and variables for
'   creating a new set of signal/waveforms
'
Sub ClearWaves()
    
    ReDim Wave(0)
    ReDim Wave(0).Sample(0)
    ReDim WrkWave.Sample(0)
    ReDim OutWave.Sample(0)
    
    ReDim PrtSize(0)
    ReDim PrtName(0)
    
    PrtSize(0) = 0
    PrtName(0) = vbNullString
    PrtNumb = 0
    PrtIndx = 0
    
    WAV.Size = 0
    ReDim WAV.Data(0)
    WAV.Data(0) = 0
    
    DoEvents
    
End Sub
' Envelope: Applies an envelope to the selected waveform
'   in order to modulate the signal level for more dynamic
'   sound presentations.
'
Sub Envelope(Which%, ByVal AttackMS&, ByVal PeakDB&, ByVal DecayMS&, ByVal SustainDB&, ByVal ReleaseMS&)

'   Which = selected waveform ID
'   AttackMS = the rise time from off to PeakDB
'   PeakDB = the maximum signal level
'   DecayMS = the time from PeakDB to SustainDB
'   SustainDB = the primary "rest" signal level
'   ReleaseMS = the time from SustainDB to off

Dim AttenValue() As Double  ' Attenuation Envelope Array
Dim AdjustBy As Double
Dim eIndx&, vIndx&
Dim LastSam As Long

'   Note: all wave forms are preset to 4000 milliseconds (4 seconds)
'   The release timing is based on that duration backwards

    ReDim AttenValue(UBound(Wave(Which).Sample))    ' An array for attenuation values

'---------------------------------------------------------------------------------------
'   Off -> AttackMS -> PeakDB
'---------------------------------------------------------------------------------------

    LastSam = (bRate * (AttackMS / 1000))   ' Determine the sample count for the attack
    If (LastSam > UBound(Wave(Which).Sample)) Then LastSam = UBound(Wave(Which).Sample)
    AdjustBy = AttenuationValue(PeakDB)    ' Determine the peak level
    
    If (LastSam > 0) Then
        For eIndx = 0 To LastSam
            AttenValue(eIndx) = (AdjustBy * (eIndx / LastSam))
        Next
    End If
    
    LastSam = (LastSam + 1)

'---------------------------------------------------------------------------------------
'   PeakDB -> DecayMS -> SustainDB
'---------------------------------------------------------------------------------------
    
    vIndx = LastSam
    LastSam = vIndx + (bRate * (DecayMS / 1000))
        
    If (LastSam > vIndx) Then
        AdjustBy = (Abs(AttenuationValue(PeakDB) - AttenuationValue(SustainDB)) / (LastSam - vIndx))
        For eIndx = vIndx To LastSam
            If (eIndx > UBound(Wave(Which).Sample)) Then GoTo Finish
            If (Abs(PeakDB) < Abs(SustainDB)) Then
                AttenValue(eIndx) = Abs((PeakDB - (AdjustBy * Abs(eIndx - LastSam))) - AttenuationValue(SustainDB))
            Else
                AttenValue(eIndx) = Abs((PeakDB - (AdjustBy * Abs(eIndx - LastSam))) + AttenuationValue(SustainDB))
            End If
        Next
    End If
    
    LastSam = (LastSam + 1)

'---------------------------------------------------------------------------------------
'   SustainDB level carried to beginning of ReleaseMS
'---------------------------------------------------------------------------------------
    
    vIndx = LastSam
    If (ReleaseMS > 0) Then
        LastSam = (UBound(Wave(Which).Sample) - (bRate * (ReleaseMS / 1000)))
    Else
        LastSam = UBound(Wave(Which).Sample)
    End If
    
    For eIndx = vIndx To LastSam
        AttenValue(eIndx) = AttenuationValue(SustainDB)
    Next
    
'---------------------------------------------------------------------------------------
'   SustainDB -> ReleaseMS -> Off
'---------------------------------------------------------------------------------------
    
    LastSam = UBound(Wave(Which).Sample)
    
    If (ReleaseMS > 0) Then
        vIndx = LastSam - (bRate * (ReleaseMS / 1000))
    Else
        vIndx = LastSam
    End If
    
    If (LastSam > vIndx) Then
        AdjustBy = (AttenuationValue(SustainDB) / (LastSam - vIndx))
        For eIndx = vIndx To LastSam
            AttenValue(eIndx) = (AttenuationValue(SustainDB) - (AdjustBy * (eIndx - vIndx)))
        Next
    End If
    
'---------------------------------------------------------------------------------------
'   Apply the Attenuation Values
'---------------------------------------------------------------------------------------
    
Finish:

    For eIndx = 0 To UBound(Wave(Which).Sample)
        AdjustBy = (AttenValue(eIndx) / ((2 ^ bWidth) - 1))
        Wave(Which).Sample(eIndx) = (Wave(Which).Sample(eIndx) * AdjustBy)
    Next
    
End Sub
' HarmonicSeries: a DEMO routine for producing a harmonic series
'   of a given frequency at staged attenuation values for 2 seconds
'
Private Sub HarmonicSeries(ByVal Freq As Double, SetSize As Integer)

Dim hIndx%

'   Note: an "InitRiff" needs to be run first to establish the
'   sample rate and bandwidth. See the ReadMe subroutine

    If (bWidth = NotSet) Then
        InitRiff 44100, SixteenBit
    Else
        ClearWaves
    End If
    
    MakeSine hIndx, Freq, 2000
    
    For hIndx = 1 To SetSize
        
' Create the harmonic wave series
        
        MakeSine hIndx, (Freq * (hIndx + 1)), 2000
        
' Attenuate the harmonic wave by (3 * harmonic-number) decibels
        
        Attenuate hIndx, AttenuationValue(3 * hIndx)
    Next
    
    MixWaves AttenuationValue(3)

    LoadRiff OutWave.Sample

'   SaveRiff [filename]

End Sub
' InitRiff initializes the sample-rate, bandwidth and RIFF header
'   This also calls the ClearWaves routine to prep the arrays
'
Sub InitRiff(ByVal SamRate&, ByVal SamSize As BandWidth)
    
'   SamRate = Sample Rate (samples per second)
'   SamSize = Bit Resolution (8, 16, 24)
    
    bWidth = SamSize
    bRate = SamRate
    
'   Highest possible value in the bandwidth
    
    Vmax = Int((2 ^ (SamSize - 1)) - 1)

'   GROUP ID HEADER
    
    RIF.Name = "RIFF"
'   RIF.Size is calculated in the SaveRiff routine
    RIF.Wave = "WAVE"

'   FORMAT CHUNK

    FMT.Name = "fmt "
    FMT.Size = 16
    FMT.AudioFormat = 1
    FMT.Channels = 1
    FMT.SampleRate = SamRate
    FMT.ByteRate = (SamRate * (SamSize / 8))
    FMT.BlockAlign = (SamSize / 8)
    FMT.BitsPerSample = SamSize

    WAV.Name = "data"
'   WAV.Size is determined in the SaveRiff routine
'   WAV.Data is assigned through the LoadRiff routine

    ClearWaves

End Sub
' LoadRiff takes an array of LONG sample values and breaks them
'   out into a stream of bytes
'
Sub LoadRiff(WavData() As Long)

'   WavData() = an array of samples

Dim wIndx&, oIndx&
Dim Bits(3) As Double
Dim dNeed&

    oIndx = (bWidth / 8)
    dNeed = (UBound(WavData) * oIndx) + (oIndx - 1)
    
    ReDim WAV.Data(dNeed)
    
    For wIndx = 0 To UBound(WAV.Data) Step oIndx
        Select Case oIndx
        Case 1                                          '   8-bit bandwidth
            WAV.Data(wIndx) = CByte(WavData(wIndx))
        Case 2                                          '   16-bit bandwidth
            Bits(3) = Abs(WavData((wIndx / oIndx)))
            Bits(1) = Int(Bits(3) / 256)
            Bits(0) = Abs(Int(Bits(3) - (Bits(1) * 256)))
            If (WavData((wIndx / oIndx)) < 0) Then Bits(1) = (255 - Bits(1))
            WAV.Data(wIndx) = CByte(Bits(0))
            WAV.Data(wIndx + 1) = CByte(Bits(1))
        Case 3                                          '   24-bit bandwidth
            Bits(3) = Abs(WavData((wIndx / oIndx)))
            Bits(2) = Int(Bits(3) / 65536)
            Bits(1) = Abs(Int((Bits(3) - (Bits(2) * 65536)) / 256))
            Bits(0) = Abs(Int(Bits(3) - (Bits(2) * 65536) - (Bits(1) * 256)))
            If (WavData((wIndx / oIndx)) < 0) Then Bits(2) = (255 - Bits(2))
            WAV.Data(wIndx) = CByte(Bits(0))
            WAV.Data(wIndx + 1) = CByte(Bits(1))
            WAV.Data(wIndx + 2) = CByte(Bits(2))
        End Select
    Next
    
'   The data size
    
    WAV.Size = (UBound(WAV.Data) + 1)

End Sub
Sub MakeSilence(ByVal MS As Long)

Dim SamCount&
    
    SamCount = (bRate * (MS / 1000))

    OutWave.Count = SamCount
    ReDim OutWave.Sample(SamCount - 1)

    For OutWave.Index = 0 To (OutWave.Count - 1)
        OutWave.Sample(OutWave.Index) = Vmax
    Next

End Sub
' MakeSine calculates sine wave values against the sample rate and bandwidth
'   given the frequency and duration of the signal in MilliSeconds
'
Function MakeSine(Which%, ByVal Freq As Double, ByVal MS As Long, Optional ByVal PhaseAngle As Double) As Boolean

'   Which = identifies waveform array (use sequentially starting with ZERO)
'   Freq = Frequency; cycles per second
'   MS = Milliseconds in duration

'   Returns TRUE if completed

Dim FreqCoeff As Double
Dim PhaseAlign As Double            ' Phase Align
Dim PhaseShift As Double            ' Phase Shift
Dim SamCount&

    SamCount = (bRate * (MS / 1000))

'   NOT FINISHED BECAUSE ...

    PhaseShift = 0

'   ... I NEED TO CALCULATE SAMPLE-OFFSET FOR THE PhaseAngle:
'   THE RELATIONSHIP BETWEEN PhaseShift AND PhaseAngle
'   It needs to be a sample-count value relative to the angle
    
    If (Which > UBound(Wave)) Then ReDim Preserve Wave(Which)

    Wave(Which).Count = 0
    Wave(Which).Index = 0
    ReDim Wave(Which).Sample(0)
    
'   If the sample count is too small, we reject it
    
    If (SamCount < 10) Then Exit Function
    
'   I NEED TO DETERMINE A MAXIMUM SAMPLE COUNT BEFORE
'   RUNNING OUT OF MEMORY
    
    FreqCoeff = (2 * (4 * Atn(1)) * (Freq / bRate))
    Wave(Which).Count = SamCount
    
    For Wave(Which).Index = 0 To (SamCount - 1)
        ReDim Preserve Wave(Which).Sample(Wave(Which).Index)
        PhaseAlign = (Wave(Which).Index + PhaseShift)
        Wave(Which).Sample(Wave(Which).Index) = 0 - (Vmax * Sin(FreqCoeff * PhaseAlign))
    Next

    MakeSine = True

End Function


' MakeSineMod calculates modulated sine wave values against the sample
'   rate and bandwidth given the frequency and duration of the signal in
'   MilliSeconds and adjusts the output by a modulation frequency and
'   amplitude.
'
Function MakeSineMod(Which%, ByVal Freq As Double, ByVal MS As Long, ByVal ModFreq As Double, ByVal ModAmp As Double) As Boolean

'   Which = identifies waveform array (use sequentially starting with ZERO)
'   Freq = Frequency; cycles per second
'   MS = Milliseconds in duration
'   ModFreq = Modulation Frequency
'   ModAmp = Modulation Amplitude

'   Returns TRUE if completed

Dim FreqCoeff As Double
Dim FreqShift As Double
Dim PhaseAlign As Double            ' Phase Align
Dim PhaseShift As Double            ' Phase Shift
Dim SamCount&

    SamCount = (bRate * (MS / 1000))

    If (Which > UBound(Wave)) Then ReDim Preserve Wave(Which)

    Wave(Which).Count = 0
    Wave(Which).Index = 0
    ReDim Wave(Which).Sample(0)
    
'   If the sample count is too small, we reject it
    
    If (SamCount < 10) Then Exit Function
    
'   I NEED TO DETERMINE A MAXIMUM SAMPLE COUNT BEFORE
'   RUNNING OUT OF MEMORY
    
    FreqCoeff = (2 * (4 * Atn(1)) * (Freq / bRate))
    FreqShift = (2 * (4 * Atn(1)) * (ModFreq / bRate))
    
    Wave(Which).Count = SamCount
    
    For Wave(Which).Index = 0 To (SamCount - 1)
        ReDim Preserve Wave(Which).Sample(Wave(Which).Index)
        
' CALCULATE THE PhaseShift BASED ON THE ModFreq AND ModAmp VALUES
' AND APPLY TO AN ADJUSTMENT AGAINST PhaseAligh
        
        PhaseShift = (ModAmp * Sin(FreqShift * Wave(Which).Index))
        
        PhaseAlign = (Wave(Which).Index + PhaseShift)
        Wave(Which).Sample(Wave(Which).Index) = 0 - (Vmax * Sin(FreqCoeff * PhaseAlign))
    Next

    MakeSineMod = True

End Function
' MixWaves will gather all of the waveforms generated by MakeSine
'   and mix them into a single stream reduced to within the
'   normalization peak value (in bits)
'
Sub MixWaves(ByVal Peak As Double)

'   Peak = maximum amplitude in bits (no higher than Vmax)
'   Use (and see) the AttenuationValue function to get the peak value

Dim Adjust As Double, WorkData As Long
Dim MaxVal As Double, MinVal As Double
Dim wIndx%, wDivis As Double

    WrkWave.Count = 0
    WrkWave.Index = 0
    ReDim WrkWave.Sample(0)

    MinVal = 0: MaxVal = 0

'   [1] Get the sample count

    For wIndx = 0 To UBound(Wave)
        If (WrkWave.Count < Wave(wIndx).Count) Then WrkWave.Count = Wave(wIndx).Count
    Next
    
    If WrkWave.Count < 100 Then Exit Sub
    ReDim WrkWave.Sample(WrkWave.Count - 1)

'   Mixing the waves together consists primarily of averaging the values

    wDivis = (UBound(Wave) + 1)

'   [2] Add the wave values together at the same strength
        
    For WrkWave.Index = 0 To (WrkWave.Count - 1)
        WrkWave.Sample(WrkWave.Index) = 0
        
'           ... even if one signal runs out - it keeps combining the values
'               and assumes the one that run out has a signal value of zero
        
        For wIndx = 0 To UBound(Wave)
            If (wIndx <= UBound(Wave(wIndx).Sample)) Then _
                WrkWave.Sample(WrkWave.Index) = (WrkWave.Sample(WrkWave.Index) + Wave(wIndx).Sample(WrkWave.Index))
        Next
        
'           divide by the number of waves being added
        
        WrkWave.Sample(WrkWave.Index) = (WrkWave.Sample(WrkWave.Index) / wDivis)

'   [3] Determine the Min and Max Normalizing Values at the same time
    
        If (WrkWave.Sample(WrkWave.Index) > MaxVal) Then MaxVal = WrkWave.Sample(WrkWave.Index)
        If ((WrkWave.Sample(WrkWave.Index) * -1) > MinVal) Then MinVal = (WrkWave.Sample(WrkWave.Index) * -1)
    Next

'   [4] Establish the normalizing value

    Adjust = 1
    
    If (MaxVal > MinVal) Then
        If (MaxVal > 0) Then Adjust = ((Peak * 0.5) / MaxVal)
    Else
        If (MinVal > 0) Then Adjust = ((Peak * 0.5) / MinVal)
    End If
    
'   [5] Apply the normalizing value
    
    For WrkWave.Index = 0 To (WrkWave.Count - 1)
        WrkWave.Sample(WrkWave.Index) = (WrkWave.Sample(WrkWave.Index) * Adjust)
    Next

'   [6] Align 8-bit samples to monopolar output.

    OutWave.Count = WrkWave.Count
    ReDim OutWave.Sample(OutWave.Count - 1)

    For OutWave.Index = 0 To (OutWave.Count - 1)
        If (bWidth = EightBit) Then             '8-bit samples are not bipolar
            WorkData = Int(WrkWave.Sample(OutWave.Index) + (Vmax + 1))
        Else                                    ' 16 and 24 bit samples are
            WorkData = Int(WrkWave.Sample(OutWave.Index))
        End If
        OutWave.Sample(OutWave.Index) = WorkData
    Next

End Sub
Function NextSine(ByVal Freq As Double, ByVal MS As Long, Optional ByVal PhaseAngle As Double) As Integer

Dim nIndx%

    If (UBound(Wave) = 0) And (Wave(0).Count = 0) Then
        nIndx = 0
    Else
        nIndx = (UBound(Wave) + 1)
    End If
    
    MakeSine nIndx, Freq, MS, PhaseAngle
    NextSine = UBound(Wave)

End Function
Function NextSineMod(ByVal Freq As Double, ByVal MS As Long, ByVal ModFreq As Double, ByVal ModAmp As Double) As Integer

Dim nIndx%

    If (UBound(Wave) = 0) And (Wave(0).Count = 0) Then
        nIndx = 0
    Else
        nIndx = (UBound(Wave) + 1)
    End If
    
    MakeSineMod nIndx, Freq, MS, ModFreq, ModAmp
    NextSineMod = UBound(Wave)

End Function
Private Sub ReadMe()

' See also the DEMO subroutine HarmonicSeries

'--------------------------------------------------------------------

'   Fundamental Application:

' Initialize the RIFF variables and buffers for a given sample rate
'   and band width:

'   InitRiff SampleRate, BandWidth

' Build the SINE wave collection:

'   MakeSine 0, Frequency0, Duration, PhaseAngle
'   MakeSine 1, Frequency1, Duration, PhaseAngle
'
'           ... etc ...
'
'   MakeSine N, FrequencyN, Duration, PhaseAngle

' Mix the waves together:

'   MixWaves AttenuationValue(X)

' Save the data:

'   LoadRiff OutWave.Sample
'   SaveRiff "FileName.wav"

' Reinitialize the variables for another wave using the same
'   sample rate and band width:

'   ClearWaves

'--------------------------------------------------------------------

' Prior to mixing the waves, they can be attenuated independently of the mix:

'   MakeSine 0, Frequency0, Duration, PhaseAngle

'       ... produces a sine wave at maximum saturation

'   Attenuate 0, AttenuationValue(3)

'       ... will adjust Wave(0) by -3db

'   MixWaves AttenuationValue(3)

'       ... will apply an addition -3db level reduction

'--------------------------------------------------------------------

'   StashChunk Usage: Permits the creation of large waves whose
'   size would exceed the memory resources of a given machine

'   ... build part into PartBuffer

'   LoadRiff OutWave.Sample
'   StashChunk

'   ... build next part

'   LoadRiff OutWave.Sample
'   StashChunk

'   SaveRiff "FileName.wav"

'--------------------------------------------------------------------

End Sub
' SaveRiff will write the PCM file to disk. (See also StashChunk)
'
Sub SaveRiff(SaveName$)

'   SaveName = file name, including path and ".wav" extension

Dim RifIndx%, pIndx%
Dim PrtStrg$
    
'   Accumulate the fragment sizes IF StashChunk was used
    
    If (PrtNumb > 0) Then
        WAV.Size = 0
        For pIndx = 0 To UBound(PrtSize)
            WAV.Size = (WAV.Size + PrtSize(pIndx))
        Next
    End If
    
'   The total size of the wave data
    
    RIF.Size = (4 + (8 + FMT.Size) + (8 + WAV.Size))
    
'   Eliminate any existing file with the same name ...
    
    RifName = SaveName
    If (RifName > vbNullString) Then
        If FileExist(RifName) Then Kill RifName
    Else
        MsgBox "No File Name"       ' ... if, of course, there is one
        End
    End If
    
    RifIndx = FreeFile
    Open RifName For Binary As RifIndx
    
'   RIF CHUNK
    
    Put #RifIndx, , RIF.Name
    Put #RifIndx, , RIF.Size
    Put #RifIndx, , RIF.Wave
    
'   FORMAT CHUNK
    
    Put #RifIndx, , FMT.Name
    Put #RifIndx, , FMT.Size
    Put #RifIndx, , FMT.AudioFormat
    Put #RifIndx, , FMT.Channels
    Put #RifIndx, , FMT.SampleRate
    Put #RifIndx, , FMT.ByteRate
    Put #RifIndx, , FMT.BlockAlign
    Put #RifIndx, , FMT.BitsPerSample
    
'   DATA CHUNK

    Put #RifIndx, , WAV.Name
    Put #RifIndx, , WAV.Size
    
'   If StashChunk had been used, we need to read and write
'   each fragment contiguously
    
    If (PrtNumb = 0) Then
        Put #RifIndx, , WAV.Data
    Else
        For pIndx = 0 To UBound(PrtName)
            PrtStrg = Space(PrtSize(pIndx))
            PrtIndx = FreeFile
            Open PrtName(pIndx) For Binary As PrtIndx
            Get #PrtIndx, , PrtStrg
            Close #PrtIndx
            Put #RifIndx, , PrtStrg
            DoEvents
            Kill PrtName(pIndx)
        Next
    End If
    Close #RifIndx

'   All Done!

End Sub
' StashChunk - in the event of a large sample, StashChunk allows
'   saving parts of the whole wave data into temporary fragments
'   which are accumulated later by the SaveRiff routine
'
Sub StashChunk()

    ReDim Preserve PrtName(PrtNumb)
    ReDim Preserve PrtSize(PrtNumb)
    
    PrtName(PrtNumb) = (App.Path & "\CHUNK." & Format(PrtNumb, "000"))

    If (Dir(PrtName(PrtNumb)) > vbNullString) Then Kill PrtName(PrtNumb)
    
    PrtIndx = FreeFile
    Open PrtName(PrtNumb) For Binary As PrtIndx
    Put #PrtIndx, , WAV.Data
    Close #PrtIndx

    PrtSize(PrtNumb) = WAV.Size

    PrtNumb = (PrtNumb) + 1

End Sub
