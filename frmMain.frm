VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BAC% Calculator - By Sc00bz"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEffectsOfAlcohol 
      Height          =   3135
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Text            =   "frmMain.frx":0000
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton cmdDrunkest 
      Caption         =   "Drunkest I've Ever Been"
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmdMeNow 
      Caption         =   "Me Now"
      Height          =   315
      Left            =   1680
      TabIndex        =   12
      Top             =   2280
      Width           =   795
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate BAC%"
      Default         =   -1  'True
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   10
      Top             =   1860
      Width           =   675
   End
   Begin VB.TextBox txtProof 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   8
      Top             =   1500
      Width           =   675
   End
   Begin VB.TextBox txtOz 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   6
      Top             =   1140
      Width           =   675
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   4
      Top             =   780
      Width           =   675
   End
   Begin VB.Frame fraGender 
      Caption         =   "Gender"
      Height          =   615
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   2235
      Begin VB.OptionButton optFemale 
         Caption         =   "Female"
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optMale 
         Caption         =   "Male"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Label lblBAC 
      Alignment       =   2  'Center
      Caption         =   "BAC%"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblTime 
      Caption         =   "Time (hr):"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label lblProof 
      Caption         =   "Proof:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label lblOz 
      Caption         =   "Amount of Alcohol (oz):"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Label lblWeight 
      Caption         =   "Weight (lbs):"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   915
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Info from:     http://www.ou.edu/oupd/bac.htm
'Formula from:  http://www.lastcall.org/topics/bac.htm
Private MouseDown As Boolean
Private Function BAC(Gender_Male As Boolean, Weight_lbs As Double, Alcohol_oz As Double, Alcohol_Proof As Double, Time_hr As Double) As Double
    'This is the long way of doing it
    'This goes through each step
    Dim Weight_kg As Double, Water_ml As Double, Alcohol_g As Double
    Dim WAlcohol_gPerml As Double, BAlcohol_gPerml As Double, Metabolism As Double
    
    If Alcohol_Proof > 200 Then Alcohol_Proof = 200
    If Weight_lbs = 0 Then Exit Function
    'convert pounds (lbs) to kilograms (kg)
    Weight_kg = Weight_lbs / 2.2046 'lbs -> kg
    'calculate total body water or water as a percent of body weight
    '1 kg of water = 1000 milliliters of water
    If Gender_Male Then
        Water_ml = 580 * Weight_kg 'men exhibit 58% water weight
    Else
        Water_ml = 490 * Weight_kg 'women exhibit 49% water weight
    End If
    'determine the weight in grams of one ounces (1 oz.) of alcohol
    '(alcohol has a specific gravity of .79)
    'x grams = 29.57ml/oz * .79 g/ml
    Alcohol_g = 29.57 * 0.79
    'calculate alcohol concentration in water
    WAlcohol_gPerml = Alcohol_g / Water_ml 'grams alcohol/ml of water
    'calculate alcohol concentration in blood (blood is composed of 80.6% water)
    BAlcohol_gPerml = WAlcohol_gPerml * 0.806 'grams alcohol/ml of blood
    'convert to standard measure (i.e. In the United States the typical measure is
    '    grams of ethanol in 100 milliliters of blood or in 210 liters of breath.)
    BAlcohol_gPerml = BAlcohol_gPerml * 100 'grams alcohol per 100 ml, or .048 BAC
    'ounces of alcohol * BAC of 1oz
    BAC = Alcohol_oz * Alcohol_Proof / 200 * BAlcohol_gPerml
    Metabolism = 0.012 * Time_hr
    BAC = BAC - Metabolism
    If BAC < 0 Then BAC = 0
End Function
Private Function BAC_Short(Gender_Male As Boolean, Weight_lbs As Double, Alcohol_oz As Double, Alcohol_Proof As Double, Time_hr As Double) As Double
    'This is the shorter/faster version
    'Most of the calculations have already been done in the BAC_CONST
    Dim Water_ml As Double
    Const BAC_CONST As Double = 0.02075454730414
    
    If Alcohol_Proof > 200 Then Alcohol_Proof = 200
    If Weight_lbs = 0 Then Exit Function
    'Calculate total body water or water as a percent of body weight
    If Gender_Male Then
        Water_ml = 0.58 * Weight_lbs 'men exhibit 58% water weight
    Else
        Water_ml = 0.49 * Weight_lbs 'women exhibit 49% water weight
    End If
    BAC_Short = Alcohol_oz * Alcohol_Proof / Water_ml * BAC_CONST - 0.012 * Time_hr
    If BAC_Short < 0 Then BAC_Short = 0
End Function
Private Function EffectsOfAlcohol(BacLevel As Double) As String
    'Info from: http://www.ou.edu/oupd/bac.htm
    Select Case CLng(Int(100# * BacLevel))
        Case Is < 2&:  EffectsOfAlcohol = "0.00 - 0.01" & vbNewLine & "Pshh drink up."
        Case Is < 4&:  EffectsOfAlcohol = "0.02 - 0.03" & vbNewLine & "No loss of coordination, slight euphoria and loss of shyness. Depressant effects are not apparent. Mildly relaxed and maybe a little lightheaded."
        Case Is < 7&:  EffectsOfAlcohol = "0.04 - 0.06" & vbNewLine & "Feeling of well-being, relaxation, lower inhibitions, sensation of warmth. Euphoria. Some minor impairment of reasoning and memory, lowering of caution. Your behavior may become exaggerated and emotions intensified (Good emotions are better, bad emotions are worse)"
        Case Is < 10&: EffectsOfAlcohol = "0.07 - 0.09" & vbNewLine & "Slight impairment of balance, speech, vision, reaction time, and hearing. Euphoria. Judgment and self-control are reduced, and caution, reason and memory are impaired (in all states .08 is legally impaired and it is illegal to drive at this level). You will probably believe that you are functioning better than you really are."
        Case Is < 13&: EffectsOfAlcohol = "0.10 - 0.125" & vbNewLine & "Significant impairment of motor coordination and loss of good judgment. Speech may be slurred; balance, vision, reaction time and hearing will be impaired. Euphoria. It is illegal to operate a motor vehicle at this level of intoxication in all states."
        Case Is < 16&: EffectsOfAlcohol = "0.13 - 0.15" & vbNewLine & "Gross motor impairment and lack of physical control. Blurred vision and major loss of balance. Euphoria is reduced and *dysphoria is beginning to appear. Judgment and perception are severely impaired. (*Dysphoria: An emotional state of anxiety, depression, or unease.)"
        Case Is < 20&: EffectsOfAlcohol = "0.16 - 0.19" & vbNewLine & "Dysphoria predominates, nausea may appear. The drinker has the appearance of a ""sloppy drunk."""
        Case Is < 25&: EffectsOfAlcohol = "0.20" & vbNewLine & "Feeling dazed/confused or otherwise disoriented. May need help to stand/walk. If you injure yourself you may not feel the pain. Some people have nausea and vomiting at this level. The gag reflex is impaired and you can choke if you do vomit. Blackouts are likely at this level so you may not remember what has happened."
        Case Is < 30&: EffectsOfAlcohol = "0.25" & vbNewLine & " All mental, physical and sensory functions are severely impaired. Increased risk of asphyxiation from choking on vomit and of seriously injuring yourself by falls or other accidents."
        Case Is < 35&: EffectsOfAlcohol = "0.30" & vbNewLine & "STUPOR. You have little comprehension of where you are. You may pass out suddenly and be difficult to awaken."
        Case Is < 40&: EffectsOfAlcohol = "0.35" & vbNewLine & "Coma is possible. This is the level of surgical anesthesia."
        Case Else:     EffectsOfAlcohol = "0.40 and up" & vbNewLine & "Onset of coma, and possible death due to respiratory arrest. (50% of the population will die at this level)"
    End Select
End Function
Private Sub txtNum_Change(txtNum As TextBox)
    Dim A As Long, B As Integer, Temp As Long
    Dim Data As String, Pos As Long, NoDot As Boolean
    
    Data = txtNum.Text
    Temp = txtNum.SelStart
    Pos = 1
    NoDot = True
    For A = 1 To Len(Data)
        B = Asc(Mid$(Data, A, 1))
        If B >= vbKey0 And B <= vbKey9 Then
            If Pos <> A Then Mid$(Data, Pos, 1) = Mid$(Data, A, 1)
            Pos = Pos + 1
        ElseIf B = 46 And NoDot Then 'Asc(".") = 46
            If Pos <> A Then Mid$(Data, Pos, 1) = Mid$(Data, A, 1)
            Pos = Pos + 1
            NoDot = False
        ElseIf A < Temp Then
            Temp = Temp - 1
        End If
    Next
    If Len(txtNum.Text) <> Pos - 1 Then
        txtNum.Text = Left$(Data, Pos - 1)
        txtNum.SelStart = Temp
    End If
End Sub
Private Sub txtNum_GotFocus(txtNum As TextBox)
    If MouseDown Then
        MouseDown = False
        Exit Sub
    End If
    txtNum.SelStart = 0
    txtNum.SelLength = Len(txtNum.Text)
End Sub
Private Sub txtNum_KeyPress(txtNum As TextBox, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, 8, 3, 24, 26
            '0-9, backspace, copy, cut, undo
        Case 1 'select all
            txtNum.SelStart = 0
            txtNum.SelLength = Len(txtNum.Text)
            KeyAscii = 0
        Case 22 'paste
            Dim A As Long, B As Integer, Data As String, Pos As Long
            Pos = 1
            Data = Clipboard.GetText
            For A = 1 To Len(Data)
                B = Asc(Mid$(Data, A, 1))
                If B >= vbKey0 And B <= vbKey9 Then
                    If Pos <> A Then Mid$(Data, Pos, 1) = Mid$(Data, A, 1)
                    Pos = Pos + 1
                End If
            Next
            Data = Left$(Data, Pos - 1)
            txtNum.SelText = Data
            KeyAscii = 0
        Case 46
            If InStr(txtNum.Text, ".") <> 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub cmdCalc_Click()
    Dim A As Long, Proof As Double, BacLevel As Double, BacStr As String
    
    Proof = Val(txtProof.Text)
    BacLevel = BAC_Short(optMale.Value, Val(txtWeight.Text), Val(txtOz.Text), Proof, Val(txtTime.Text))
    txtEffectsOfAlcohol.Text = EffectsOfAlcohol(BacLevel)
    
    BacStr = CStr(Round(BacLevel, 4))
    A = InStr(BacStr, ".")
    If A = 0 Then
        BacStr = BacStr & ".0000"
    ElseIf 4& - Len(BacStr) + A > 0 Then
        BacStr = BacStr & String$(4& - Len(BacStr) + A, "0")
    End If
    
    lblBAC.Caption = Round(BacLevel, 4) & " BAC%"
    txtWeight.Text = Val(txtWeight.Text)
    txtOz.Text = Val(txtOz.Text)
    txtProof.Text = Proof
    txtTime.Text = Val(txtTime.Text)
End Sub
Private Sub cmdDrunkest_Click()
'I was somewhere between
'0.30: STUPOR. You have little comprehension of where you are.
'      You may pass out suddenly and be difficult to awaken.
'0.35: Coma is possible. This is the level of surgical anesthesia.

'                     The Drunkest I've Ever Been
'This was a fun time. It started out as a normal Wednesday night until
'I wanted to have a shot of Spirytus with this kid on my floor. We did
'and his reaction was awesome he like chugged a bottle of water and ran
'to the window and started coughing. Well he stopped drinking but I was
'like hmm I have class at 9:30am I'll have another. One turned into two,
'two into three... yeah. Then I knew I was fucked up but I was like hmm
'maybe if I do half shots I won't get more fucked up. After a while, I
'was on my way to passing out and someone was like Sc00bz wake up don't
'fall asleep. I'm all like wtf I'm tired. Then the kid I was drinking
'with and someone else from my floor picked me up and took me to the
'bathroom. I drink sooooo much water and pissed so much. About 2-3 hrs
'later it's time for bed. I walk into my room and my roommate, in his
'sleep, mumbles something about defecating on a TV or computer or
'something. I stormed out of there laughing my ass off. I got half way
'down the hall before collapsing. The next day I had an upset stomach
'and no hang over but I did feel like throwing up for the next two days.
'Then the weekend came... ;)

    optMale.Value = True
    txtWeight.Text = "155"  '155 lbs
    txtOz.Text = "8.454"    '250ml (1/3 of a bottle of Spirytus straight)
    txtProof.Text = "192"   '96%
    txtTime.Text = "2.5"    '2 hrs 30 min
    cmdCalc_Click
End Sub
Private Sub cmdMeNow_Click()
    optMale.Value = True
    txtWeight.Text = "155"  '155 lbs
    txtOz.Text = "6"        '3 shooters about 2oz each
    txtProof.Text = "84.67" '40% (Jack Daniel's Whiskey)
                            '40% (Stolichnaya Russian Vodka)
                            '47% (Beefeater Dry Gin)
                            '42.33% ave
    txtTime.Text = "1.5"    '1 hr 30 min
    cmdCalc_Click
End Sub
Private Sub txtEffectsOfAlcohol_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 0, 8, 22, 28 To 255 '1 to 26 are Ctrl+? 27 is Escape
            '?, Backspace, Paste, Non Ctrl+?
            MsgBox "Editing this text box will not change the effects of alcohol in you.", vbInformation, "The Effects of Alcohol in you"
        Case 1 'Select all
            txtEffectsOfAlcohol.SelStart = 0
            txtEffectsOfAlcohol.SelLength = Len(txtEffectsOfAlcohol.Text)
            KeyAscii = 0
        Case 19 'Ctrl+S (Save)
            MsgBox "No one can save you now, you're too drunk.", vbInformation, "The Effects of Alcohol in you"
        Case 3, 24 'Ctrl+C (Copy), Ctrl+X (Cut)
            'When a text box is locked you need to do code for copying
            Clipboard.Clear
            Clipboard.SetText txtEffectsOfAlcohol.SelText
        Case 26 'Ctrl+Z (Undo)
            MsgBox "You can't undo the effects of alcohol in you." & vbNewLine & "There is no Ctrl+Z in life or someone would of done it to you a long time ago.", vbInformation, "The Effects of Alcohol in you"
    End Select
End Sub
'Everything after this comment is for making the
'input text boxes numbers only with one decimal.
'Also for when you use tab it has every thing selected
Private Sub txtOz_Change()
    txtNum_Change txtOz
End Sub
Private Sub txtOz_GotFocus()
    txtNum_GotFocus txtOz
End Sub
Private Sub txtOz_KeyPress(KeyAscii As Integer)
    txtNum_KeyPress txtOz, KeyAscii
End Sub
Private Sub txtOz_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
End Sub
Private Sub txtOz_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub
Private Sub txtProof_Change()
    txtNum_Change txtProof
End Sub
Private Sub txtProof_GotFocus()
    txtNum_GotFocus txtProof
End Sub
Private Sub txtProof_KeyPress(KeyAscii As Integer)
    txtNum_KeyPress txtProof, KeyAscii
End Sub
Private Sub txtProof_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
End Sub
Private Sub txtProof_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub
Private Sub txtTime_Change()
    txtNum_Change txtTime
End Sub
Private Sub txtTime_GotFocus()
    txtNum_GotFocus txtTime
End Sub
Private Sub txtTime_KeyPress(KeyAscii As Integer)
    txtNum_KeyPress txtTime, KeyAscii
End Sub
Private Sub txtTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
End Sub
Private Sub txtTime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub
Private Sub txtWeight_Change()
    txtNum_Change txtWeight
End Sub
Private Sub txtWeight_GotFocus()
    txtNum_GotFocus txtWeight
End Sub
Private Sub txtWeight_KeyPress(KeyAscii As Integer)
    txtNum_KeyPress txtWeight, KeyAscii
End Sub
Private Sub txtWeight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = True
End Sub
Private Sub txtWeight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown = False
End Sub
