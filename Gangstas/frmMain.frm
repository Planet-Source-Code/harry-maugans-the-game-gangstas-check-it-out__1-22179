VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gangsta'"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   570
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   4680
      Top             =   0
   End
   Begin VB.Label lblOP 
      Caption         =   "No Opponent"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblPower 
      Caption         =   "Your Power: 2"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblOH 
      Caption         =   "No Opponent"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblOpponent 
      Caption         =   "No Opponent"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblHealth 
      Caption         =   "Your Health: 10"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Label lblWeapon 
      Caption         =   "Weapon: Butter Knife"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label lblVehicle 
      Caption         =   "Vehicle: Tricycle"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblMoney 
      Caption         =   "Money: 50"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblBalls 
      Caption         =   "Balls: 0"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   5460
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   6180
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuGameLoad 
         Caption         =   "&Load"
      End
   End
   Begin VB.Menu mnuBank 
      Caption         =   "&Bank"
      Begin VB.Menu mnuBankWith 
         Caption         =   "&Withdraw"
      End
      Begin VB.Menu mnuBankDeposit 
         Caption         =   "&Deposit"
      End
      Begin VB.Menu mnuBankSum 
         Caption         =   "&Summary"
      End
   End
   Begin VB.Menu mnuTrade 
      Caption         =   "&Trade"
      Begin VB.Menu mnuTradeWeapons 
         Caption         =   "&Weapons"
      End
      Begin VB.Menu mnuTradeVehicles 
         Caption         =   "&Vehicles"
      End
   End
   Begin VB.Menu mnuFight 
      Caption         =   "&Fight"
      Begin VB.Menu mnuFightQuickieMart 
         Caption         =   "&Quickie Mart"
      End
      Begin VB.Menu mnuFightStore 
         Caption         =   "&General Store"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFightWerehouse 
         Caption         =   "&Werehouse"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFightBank 
         Caption         =   "&Bank"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDo 
      Caption         =   "&Do"
      Begin VB.Menu mnuDoDrink 
         Caption         =   "&Drink"
      End
      Begin VB.Menu mnuDoSmoke 
         Caption         =   "&Smoke"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Money As Integer
Public Vehicle As String
Public Balls As Integer
Public Weapon As String
Public Power As Integer
Public Health As Integer
Public MaxHealth As Integer
Public EnemyHealth As Integer
Public BankMoney As Integer
Public Opponent As String
Public OH As Integer
Public OP As Integer
Public Area As String

Private Sub Form_Load()
    Money = 50
    Vehicle = "Tricycle"
    Balls = 0
    Weapon = "Butter Knife"
    Power = 2
    Health = 10
    MaxHealth = 10
    EnemyHealth = 0
    DoEvents
    MsgBox "Win by having 100 balls."
End Sub

Private Sub mnuBankDeposit_Click()
    Dim value As String
    
    If Money < 1 Then MsgBox "You don't have any money to deposit!": Exit Sub
    value = InputBox("Please enter an ammount to deposit from 1 to " & Money, "Deposit")
    If IsNumeric(value) = False Then MsgBox "Please enter a valaid number.": Exit Sub
    If value > Money Then MsgBox "You don't have that much money!": Exit Sub
    Money = Money - value
    BankMoney = BankMoney + value
    lblMoney.Caption = "Money: " & Money
End Sub

Private Sub mnuBankSum_Click()
    MsgBox "There is " & BankMoney & " in the bank." & vbCrLf & "You have " & Money & " with you.", , "Bank Summary"
End Sub

Private Sub mnuBankWith_Click()
    Dim value As String
    
    If BankMoney < 1 Then MsgBox "No money in the bank!": GoTo gt
    value = InputBox("Please enter an ammount to withdraw from 1 to " & BankMoney, "Withdraw")
    If IsNumeric(value) = False Then MsgBox "Please enter a valaid number.": GoTo gt
    If value > BankMoney Then MsgBox "You don't have that much money in the bank!": GoTo gt
    BankMoney = BankMoney - value
    Money = Money + value
    lblMoney.Caption = "Money: " & Money
gt:
    CheckforRobbers
End Sub

Private Sub mnuDoDrink_Click()
    Dim result As String
    
    result = MsgBox("A beer costs 10 bucks.  You up with that?", vbYesNo, "Beer...")
    If result = vbNo Then GoTo ex
    If Money < 10 Then MsgBox "Not enough money!": Exit Sub
    Money = Money - 10
    MsgBox "Oooh, beer is good...make me think i am invincible...well, maybe not."
    MsgBox "Friends: You stoopid man"
    Health = Health + 1
    MaxHealth = MaxHealth + 1
    Balls = Balls - 1
    lblMoney.Caption = "Money: " & Money
    lblBalls.Caption = "Balls: " & Balls
    lblHealth.Caption = "Health: " & MaxHealth
ex:
    CheckforRobbers
End Sub

Private Sub mnuDoSmoke_Click()
    Dim result As String
    
    result = MsgBox("A joint costs 100 bucks.  You up with that?", vbYesNo, "Joint")
    If result = vbNo Then GoTo ex
    If Money < 100 Then MsgBox "Not enough money!": Exit Sub
    Money = Money - 100
    MsgBox "Inhale....puff..inhale...puff..now dat shit is strong...wow, look at the little birdy..."
    MsgBox "Friends: Go dude!"
    Balls = Balls + 5
    lblBalls.Caption = "Balls: " & Balls
    lblMoney.Caption = "Money: " & Money
    CheckBalls
ex:
    CheckforRobbers
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFightBank_Click()
    Opponent = "Bank Vault"
    OH = 300
    OP = 15
    Me.Enabled = False
    lblOpponent.Caption = "Bank Vault"
    Area = "BV"
    lblOH.Caption = "300"
    Timer1.Enabled = True
End Sub

Private Sub mnuFightQuickieMart_Click()
    Randomize
    Opponent = Int((3 * Rnd) + 1)
    Select Case Opponent
    Case 1:
        Opponent = "Shop Owner"
        OH = Int((25 * Rnd) + 1)
        OP = Int((2 * Rnd) + 1)
    Case 2:
        Opponent = "Shopper"
        OH = Int((15 * Rnd) + 1)
        OP = Int((1 * Rnd) + 1)
    Case 3:
        Opponent = "Door"
        OH = Int((30 * Rnd) + 1)
        OP = Int((2 * Rnd) + 1)
    End Select
    Me.Enabled = False
    lblOpponent.Caption = Opponent
    lblOH.Caption = OH
    lblOP.Caption = "Op. Power: " & OP
    Area = "QM"
    Timer1.Enabled = True
End Sub

Private Sub mnuFightStore_Click()
    Randomize
    Opponent = Int((5 * Rnd) + 1)
    Select Case Opponent
    Case 1:
        Opponent = "Manager"
        OH = Int((50 * Rnd) + 1)
        OP = Int((5 * Rnd) + 1)
    Case 2:
        Opponent = "Owner"
        OH = Int((65 * Rnd) + 1)
        OP = Int((7 * Rnd) + 1)
    Case 3:
        Opponent = "Shopper"
        OH = Int((30 * Rnd) + 1)
        OP = Int((4 * Rnd) + 1)
    Case 4:
        Opponent = "Employee"
        OH = Int((35 * Rnd) + 1)
        OP = Int((3 * Rnd) + 1)
    Case 5:
        Opponent = "Cardboard Box"
        OH = Int((25 * Rnd) + 1)
        OP = Int((2 * Rnd) + 1)
    End Select
    Me.Enabled = False
    Area = "Store"
    lblOpponent.Caption = Opponent
    lblOP.Caption = "Op. Power: " & OP
    lblOH.Caption = OH
    Timer1.Enabled = True
End Sub

Private Sub mnuFightWerehouse_Click()
    Randomize
    Opponent = Int((5 * Rnd) + 1)
    Select Case Opponent
    Case 1:
        Opponent = "Bum"
        OH = Int((75 * Rnd) + 1)
        OP = Int((7 * Rnd) + 1)
    Case 2:
        Opponent = "Pizza Guy"
        OH = Int((60 * Rnd) + 1)
        OP = Int((6 * Rnd) + 1)
    Case 3:
        Opponent = "Spazmatic Monkey"
        OH = Int((65 * Rnd) + 1)
        OP = Int((7 * Rnd) + 1)
    Case 4:
        Opponent = "Bob"
        OH = Int((90 * Rnd) + 1)
        OP = Int((2 * Rnd) + 1)
    Case 5:
        Opponent = "Pipe Man"
        OH = Int((150 * Rnd) + 1)
        OP = Int((1 * Rnd) + 1)
    End Select
    Me.Enabled = False
    Area = "WH"
    lblOpponent.Caption = Opponent
    lblOP.Caption = "Op. Power: " & OP
    lblOH.Caption = OH
    Timer1.Enabled = True
End Sub

Private Sub mnuGameLoad_Click()
    Dim Place As String
    
    Place = InputBox("Please enter a filename to load. (ex. MySave.dat)", "Filename to load?")
    If Trim(Place) = "" Then Exit Sub
    Open App.Path & "\Data\" & Place For Input As #1
        Input #1, Money, Vehicle, Balls, Weapon, Power, Health, MaxHealth, BankMoney
    Close #1
        If Vehicle = "Tricycle" Then mnuFightStore.Visible = False: mnuFightWerehouse.Visible = False: mnuFightBank.Visible = False
        If Vehicle = "Skateboard" Then mnuFightStore.Visible = True: mnuFightWerehouse.Visible = False: mnuFightBank.Visible = False
        If Vehicle = "Bike" Then mnuFightStore.Visible = True: mnuFightWerehouse.Visible = True: mnuFightBank.Visible = False
        If Vehicle = "Car" Then mnuFightStore.Visible = True: mnuFightWerehouse.Visible = True: mnuFightBank.Visible = True
        lblMoney.Caption = "Money: " & Money
        lblVehicle.Caption = "Vehicle: " & Vehicle
        lblBalls.Caption = "Balls: " & Balls
        lblWeapon.Caption = "Weapon: " & Weapon
        lblPower.Caption = "Your Power: " & Power
        lblHealth.Caption = "Your Health: " & Health
    MsgBox "Load Successful!"
End Sub

Private Sub mnuGameSave_Click()
    Dim Place As String
    
    Place = InputBox("Please enter a filename to save to. (ex. MySave.dat)", "Filename to Save As?")
    If Trim(Place) = "" Then Exit Sub
    Open App.Path & "\Data\" & Place For Output As #1
        Write #1, Money, Vehicle, Balls, Weapon, Power, Health, MaxHealth, BankMoney
    Close #1
    MsgBox "Save Successful!"
End Sub

Private Sub mnuTradeVehicles_Click()
    Dim result As String
    
    CheckforRobbers
    If Vehicle = "Tricycle" Then
        result = MsgBox("I'll trade you my skateboard for your tricycle and 65 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 65 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 65
        Vehicle = "Skateboard"
        lblVehicle.Caption = "Vehicle: Skateboard"
        MsgBox "You can now get to the store!"
        mnuFightStore.Visible = True
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Vehicle = "Skateboard" Then
        result = MsgBox("I'll trade you my bike for your skateboard and 100 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 100 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 100
        Vehicle = "Bike"
        lblVehicle.Caption = "Vehicle: Bike"
        MsgBox "You can now get to the werehouse!"
        mnuFightWerehouse.Visible = True
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Vehicle = "Bike" Then
        result = MsgBox("I'll trade you my car for your bike and 500 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 500 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 500
        Vehicle = "Car"
        lblVehicle.Caption = "Vehicle: Car"
        MsgBox "You can now get to the bank!"
        mnuFightBank.Visible = True
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Vehicle = "Car" Then MsgBox "Why would you want more than a car?": Exit Sub
End Sub

Private Sub mnuTradeWeapons_Click()
    Dim result As String
    
    CheckforRobbers
    If Weapon = "Butter Knife" Then
        result = MsgBox("I'll trade you my cap gun for your butter knife and 75 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 75 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 75
        Weapon = "Cap Gun"
        Power = 3
        lblWeapon.Caption = "Weapon: Cap Gun"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Cap Gun" Then
        result = MsgBox("I'll trade you my razor blade for your cap gun and 80 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 80 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 80
        Weapon = "Razor Blade"
        Power = 4
        lblWeapon.Caption = "Weapon: Razor Blade"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Razor Blade" Then
        result = MsgBox("I'll trade you my knife for your razor blade and 100 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 100 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 100
        Weapon = "Knife"
        Power = 5
        lblWeapon.Caption = "Weapon: Knife"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Knife" Then
        result = MsgBox("I'll trade you my BB Gun for your knife and 115 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 115 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 115
        Weapon = "BB Gun"
        Power = 6
        lblWeapon.Caption = "Weapon: BB Gun"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "BB Gun" Then
        result = MsgBox("I'll trade you my Pistol for your BB Gun and 120 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 120 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 120
        Weapon = "Pistol"
        Power = 7
        lblWeapon.Caption = "Weapon: Pistol"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Pistol" Then
        result = MsgBox("I'll trade you my Crossbow for your Pistol and 120 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 120 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 120
        Weapon = "Crossbow"
        Power = 8
        lblWeapon.Caption = "Weapon: Crossbow"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Crossbow" Then
        result = MsgBox("I'll trade you my Hunting Rifle for your Crossbow and 125 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 125 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 125
        Weapon = "Hunting Rifle"
        Power = 9
        lblWeapon.Caption = "Weapon: Hunting Rifle"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Hunting Rifle" Then
        result = MsgBox("I'll trade you my Shotgun for your Hunting Rifle and 150 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 150 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 150
        Weapon = "Shotgun"
        Power = 10
        lblWeapon.Caption = "Weapon: Shotgun"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Shotgun" Then
        result = MsgBox("I'll trade you my bazooka for your shotgun and 250 dollars.  Wanna trade?", vbYesNo)
        If result = vbNo Then Exit Sub
        If Money < 250 Then MsgBox "You don't have enough money!": Exit Sub
        Money = Money - 250
        Weapon = "Bazooka"
        Power = 15
        lblWeapon.Caption = "Weapon: Bazooka"
        lblPower.Caption = "Your Power: " & Power
        lblMoney.Caption = "Money: " & Money
        Exit Sub
    End If
    If Weapon = "Bazooka" Then
        MsgBox "You already have the best weapon.  How could you want more than a bazooka..."
        Exit Sub
    End If
    
End Sub

Private Sub Timer1_Timer()
    If Timer1.Interval = "1500" Then
       If Health < 1 Then Dead: Exit Sub
       If OH < 1 Then Killed: Exit Sub
       Health = Health - OP
       lblHealth.Caption = "Your Health: " & Health
       OH = OH - Power
       lblOH.Caption = "Op. Health: " & OH
    End If
End Sub

Public Function Dead()
    MsgBox "Noooo....ack! The " & Opponent & " killed you!" & vbCrLf & "Good Game, you were only missing " & 100 - Balls & " balls!"
    End
End Function

Public Function Killed()
    Dim Data As String
    Dim Dollars As Integer
    
    Me.Enabled = True
    Timer1.Enabled = False
    MsgBox "Oh my god!  You killed " & Opponent & "!  You bastard!"
    MsgBox "Friends: Kool...."
    lblOpponent.Caption = "No Opponent"
    lblOH.Caption = "No Opponent"
    lblOP.Caption = "No Opponent"
    Select Case Area
    Case "QM":
        Dollars = Int((20 * Rnd) + 1)
    Case "Store":
        Dollars = Int((30 * Rnd) + 1)
    Case "WH":
        Dollars = Int((50 * Rnd) + 1)
    Case "BV":
        Dollars = 150
        MsgBox "You robbed a bank!  You friends will like this!"
        Balls = Balls + 5
        lblBalls.Caption = "Balls: " & Balls
        CheckBalls
    End Select
    MsgBox "You found " & Dollars & " dollars on the corpse!"
    Money = Money + Dollars
    lblMoney.Caption = "Money: " & Money
    Balls = Balls + 1
    lblBalls.Caption = "Balls: " & Balls
    CheckBalls
    CheckforRobbers
    Data = MsgBox("A doctor offers to repair your wounds for 2 dollars.  Do ya wanna?", vbYesNo)
    If Data = vbNo Then Exit Function
    If Money < 2 Then MsgBox "Not enough money man...": Exit Function
    Money = Money - 2
    lblMoney.Caption = "Money: " & Money
    Health = MaxHealth
    lblHealth.Caption = "Your Health: " & Health
End Function

Public Function CheckBalls()
    If Balls >= 100 Then
        MsgBox "Wow, you won!  Congrats, you friends must think you are like a hero or something..."
        MsgBox "...sirens...slam....Freeze!...You have the right to remain silent!..."
        MsgBox "Oh shit....bang, bang, you shoot.....and are shot in the leg"
        MsgBox "They drag you in and you end up getting the chair.  Life sux huh?"
        End
    End If
End Function

Public Function CheckforRobbers()
    Dim i, a As Integer
    Randomize
    a = Int((11 * Rnd) + 1)
    If a > 3 Then Exit Function
    i = Int(((Money - ((Right(Money, 1)) * 2)) * Rnd) + 1)
    If i > Money Then Exit Function
    Select Case a
    Case 1:
        MsgBox "You were mugged!  The theif took " & i & " dollars!"
    Case 2:
        MsgBox "You were pickpocketed!  You lost " & i & " dollars!"
    Case 3:
        MsgBox "You had a hole in your pocket and " & i & " dollars dropped out!"
    End Select
        Money = Money - i
        lblMoney.Caption = "Money: " & Money
End Function
