'---------------------------------------------------------------------
'THIS PROGRAM WAS MADE IN COLLABORATION OF:
'Jasmine Faye Andres
'Gene Rev Feliciano
'Angel Juganas
'---------------------------------------------------------------------

Public Class Form1
    'Declaration of Primary variables-------------------------------------
    'CHARACTERS-------------------------------------------
    Dim strCharName As String 'Character Name
    Dim intLevel As Integer 'Character Level

    'WEAPONS----------------------------------------------
    Dim strWeaponName As String 'Weapon Name
    Dim intWeapLevel As Integer 'Weapon Level
    Dim intWeapRefine As Integer 'Weapon Refinement

    'ARTIFACTS--------------------------------------------
    Dim strArtifactFlower As String 'Flower of Life
    Dim strArtifactPlume As String 'Plume of Death
    Dim strArtifactSands As String 'Sands of Eon
    Dim strArtifactGoblet As String 'Goblet of Eonothem
    Dim strArtifactCirclet As String 'Circlet of Logos

    Private Sub TableBindingNavigatorSaveItem_Click(sender As Object, e As EventArgs)
        Me.Validate()
        Me.TableBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.Database1DataSet)

    End Sub
    'Form_Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call InitializeLists()
        Call PlayBackgroundMusic()
    End Sub
    'Form_Closed
    Private Sub GICalcMainGUI_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        'SplashScreen.Close() 'Close the first Form
    End Sub
    'Music Sub
    Sub PlayBackgroundMusic()
        'My.Computer.Audio.Play("Yu Peng Chen - Reminisence-notloud.wav", AudioPlayMode.BackgroundLoop)
    End Sub
    'This sub allows us to create a mute button
    Private Sub chkMuteMusic_CheckedChanged(sender As Object, e As EventArgs) Handles chkMuteMusic.CheckedChanged
        If chkMuteMusic.Checked = True Then
            My.Computer.Audio.Stop()
        Else
            PlayBackgroundMusic()
        End If
    End Sub
    '
    Private Sub chkDisplayPresets_CheckedChanged(sender As Object, e As EventArgs) Handles chkDisplayPresets.CheckedChanged
        If chkDisplayPresets.Checked = True Then
            grpPresetBuilds.Show()
            lblPresetCreator.Show()
        Else
            grpPresetBuilds.Hide()
            lblPresetCreator.Hide()
        End If
    End Sub
    '-----------------------------------------------------------------------
    'RADIO BUTTON CONDITIONS-------------------------------------------------
    '-----------------------------------------------------------------------
    Private Sub radButtons_CheckChanged(sender As Object, e As EventArgs) Handles radWeaponStats.CheckedChanged, radPreviewStats.CheckedChanged, radCharacterStats.CheckedChanged, radArtifactStats.CheckedChanged
        grpCharStats.Hide()
        grpWeaponStats.Hide()
        grpArtifactStats.Hide()
        grpPreviewStats.Hide()
        grpDatabase.Hide()
        picCharacter.Hide() 'hide the character portrait

        If radCharacterStats.Checked = True Then
            grpCharStats.Show()
            picCharacter.Show() 'show the character portrait

        ElseIf radWeaponStats.Checked = True Then
            grpWeaponStats.Show()

        ElseIf radArtifactStats.Checked = True Then
            grpArtifactStats.Show()

        ElseIf radPreviewStats.Checked = True Then
            grpPreviewStats.Show()
            picCharacter.Show() 'show the character portrait

            Call CalculateStats()

        ElseIf radDatabase.Checked() = True Then
            grpDatabase.Show()
        End If
    End Sub
    '--------------------------------------------------------------------------------------------------
    'LIST BOX ITEM GENERATION--------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------------
    'Adds items to the lis
    Private Sub InitializeLists()
        'Characters
        lstCharacters.Items.Add("Albedo")
        lstCharacters.Items.Add("Eula")
        lstCharacters.Items.Add("Hu Tao")
        lstCharacters.Items.Add("Kaedehara Kazuha")
        lstCharacters.Items.Add("Kamisato Ayaka")
        lstCharacters.Items.Add("Kamisato Ayato")
        lstCharacters.Items.Add("Sangonomiya Kokomi")
        lstCharacters.Items.Add("Xiao")
        lstCharacters.Items.Add("Yae Miko")
        lstCharacters.Items.Add("Zhongli")
        lstCharacters.SelectedItem = "Hu Tao"

        'Character Level(s)
        lstLevel.Items.Add("1")
        For intCount As Integer = 10 To 90 Step 10
            lstLevel.Items.Add(intCount.ToString)
        Next intCount
        lstLevel.SelectedItem = "1"

        'Weapon refinement(s)
        For intCount As Integer = 1 To 5 Step 1
            lstWeaponRefinement.Items.Add(intCount.ToString)
        Next
        lstWeaponRefinement.SelectedItem = "1"

        'PRESET BUILDS LIST
        lstPresetList.Items.Add("Hu Tao")
        lstPresetList.Items.Add("Yae Miko")
        lstPresetList.Items.Add("Kamisato Ayaka")
    End Sub
    'Code that adds Weapon Levels to the item list
    Private Sub GenerateWeaponLevels()
        lstWeaponLevel.Items.Clear() 'clear the list first to prevent duplications 
        lstWeaponLevel.Items.Add("1")
        For intCount As Integer = 10 To 90 Step 10
            lstWeaponLevel.Items.Add(intCount.ToString)
        Next intCount
        lstWeaponLevel.SelectedItem = "1"
    End Sub

    'Code for lstCharacters which will run everytime a new character is selected
    Private Sub lstCharacters_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCharacters.SelectedIndexChanged
        Dim strNativeName As String = "Default" 'Displays the characters name in their own respective character systems
        Dim strWeaponType As String = "Default" 'weapon type of the character
        lstWeapons.Items.Clear() 'Clear the list first
        strCharName = Convert.ToString(lstCharacters.SelectedItem)
        Select Case strCharName
            Case "Albedo"
                picCharacter.Image = picAlbedo.Image
                strNativeName = ""
                strWeaponType = "Sword"
            Case "Hu Tao"
                'grpCharStats.BackColor = Color.LightCoral 'nagpapalit ng kulay kaso kapanget hahaha kaya wag na
                picCharacter.Image = picHuTao.Image
                strNativeName = "胡桃"
                strWeaponType = "Polearm"
            Case "Kaedehara Kazuha"
                picCharacter.Image = picKazuha.Image
                strNativeName = "楓原万葉"
                strWeaponType = "Sword"
            Case "Kamisato Ayaka"
                picCharacter.Image = picAyaka.Image
                strNativeName = "神里綾華"
                strWeaponType = "Sword"
            Case "Kamisato Ayato"
                picCharacter.Image = picAyato.Image
                strNativeName = "神里綾人"
                strWeaponType = "Sword"
            Case "Eula"
                picCharacter.Image = picEula.Image
                strNativeName = ""
                strWeaponType = "Greatsword"
            Case "Xiao"
                picCharacter.Image = picXiao.Image
                strNativeName = "魈"
                strWeaponType = "Polearm"
            Case "Sangonomiya Kokomi"
                picCharacter.Image = picKokomi.Image
                strNativeName = "珊瑚宮心海"
                strWeaponType = "Catalyst"
            Case "Yae Miko"
                picCharacter.Image = picYae.Image
                strNativeName = "八重神子"
                strWeaponType = "Catalyst"
            Case "Zhongli"
                picCharacter.Image = picZhongli.Image
                strNativeName = "钟离"
                strWeaponType = "Polearm"
        End Select
        Call AvailableWeaponSelection(strWeaponType)

        'DISPLAYING THE DISPLAYABLES :DDDD (can't think of a better way to describe it)
        lblCSCharName.Text = strCharName
        lblPSCharName.Text = strCharName
        lblCSNativeName.Text = strNativeName
        lblPSNativeName.Text = strNativeName
    End Sub

    'ONLY DISPLAYS THE APPROPRIATE WEAPONS ON THE DEPENDING CHARACTERS AND THEIR WEAPON TYPES
    'Polearm Characters - Hu Tao, Zhongli, Xiao
    'Sword Characters - Kamisato Ayaka, Kamisato Ayato, Kaedehara Kazuha, Albedo
    'Greatsword Character - Eula
    'Catalyst Characters - Sangonomiya Kokomi, Yae Miko
    Private Sub AvailableWeaponSelection(ByVal strWeaponType As String)
        Select Case strWeaponType
            Case "Polearm"
                lstWeapons.Items.Add("Beginner's Protector")
                lstWeapons.Items.Add("The Catch")
                lstWeapons.Items.Add("Staff of Homa")
                lstWeapons.Items.Add("Primordial Jade-Winged Spear")
                lstWeapons.SelectedItem = "Beginner's Protector"
            Case "Sword"
                lstWeapons.Items.Add("Amenoma Kageuchi")
                lstWeapons.Items.Add("Sacrificial Sword")
                lstWeapons.Items.Add("Mistsplitter Reforged")
                lstWeapons.SelectedItem = "Amenoma Kageuchi"
            Case "Greatsword"
                lstWeapons.Items.Add("Song of Broken Pines")
                lstWeapons.Items.Add("Wolf's Gravestone")
                lstWeapons.Items.Add("Luxurious Sea-Lord")
                lstWeapons.SelectedItem = "Song of Broken Pines"
            Case "Catalyst"
                lstWeapons.Items.Add("Everlasting Moonglow")
                lstWeapons.Items.Add("Oathsworn Eye")
                lstWeapons.Items.Add("Kagura's Verity")
                lstWeapons.SelectedItem = "Everlasting Moonglow"
        End Select
    End Sub
    'WEAPON
    Private Sub lstWeapons_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstWeapons.SelectedIndexChanged
        GenerateWeaponLevels()
        strWeaponName = Convert.ToString(lstWeapons.SelectedItem)
        Select Case strWeaponName
            Case "Beginner's Protector"
                picWeapon.Image = picBeginnersProtector.Image

                'This makes Beginner's Protector reach only up to lv 70 ONLY as it is only a 3-star weapon 
                lstWeaponLevel.Items.Clear()
                lstWeaponLevel.Items.Add("1")
                For intCount As Integer = 10 To 70 Step 10
                    lstWeaponLevel.Items.Add(intCount.ToString)
                Next intCount
                lstWeaponLevel.SelectedItem = "1"
            Case "Mistsplitter Reforged"
                picWeapon.Image = picMistsplitterReforged.Image
            Case "Staff of Homa"
                picWeapon.Image = picHoma.Image
            Case "Amenoma Kageuchi"
                picWeapon.Image = picAmenoma.Image
            Case "Sacrificial Sword"
                picWeapon.Image = picSacrificial.Image
            Case "The Catch"
                picWeapon.Image = picTheCatch.Image
            Case "Song of Broken Pines"
                picWeapon.Image = picBrokenPines.Image
            Case "Wolf's Gravestone"
                picWeapon.Image = picWolfsGravestone.Image
            Case "Everlasting Moonglow"
                picWeapon.Image = picEverlastingMoonglow.Image
            Case "Oathsworn Eye"
                picWeapon.Image = picOathswornEye.Image
            Case "Luxurious Sea-Lord"
                picWeapon.Image = picFishSword.Image
            Case "Primordial Jade-Winged Spear"
                picWeapon.Image = picPrimordialJWS.Image
            Case "Kagura's Verity"
                picWeapon.Image = picKagurasVerity.Image
        End Select

        'DISPLAY THE WEAPON NAME(s) IN THE LABELS 
        lblWeaponName.Text = strWeaponName

        Call CalculateStats() 'Update stats
    End Sub
    'CHARACTER LEVEL
    Private Sub lstLevel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstLevel.SelectedIndexChanged
        Integer.TryParse(lstLevel.SelectedItem, intLevel)
        lblCSLevel.Text = intLevel.ToString
        lblPSLevel.Text = intLevel.ToString

        Call CalculateStats() 'Update stats
    End Sub
    'WEAPON REFINEMENT
    Private Sub lstWeaponRefinement_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstWeaponRefinement.SelectedIndexChanged
        Integer.TryParse(lstWeaponRefinement.SelectedItem, intWeapRefine)
        lblWeapRefine.Text = Convert.ToString(intWeapRefine)

        Call CalculateStats() 'Update stats
    End Sub
    'WEAPON LEVEL
    Private Sub lstWeaponLevel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstWeaponLevel.SelectedIndexChanged
        Integer.TryParse(lstWeaponLevel.SelectedItem, intWeapLevel)
        lblWeaponLevel.Text = intWeapLevel.ToString

        Call CalculateStats() 'Update stats
    End Sub

    '-----------------------------------------------
    'MAIN MATH PROCESSES----------------------------
    '-----------------------------------------------
    Private Sub CalculateStats()
        'CHARACTER BASE STATS SELECTION------------------------------------------------------------------

        'generating an array index based on the character's level
        Dim intCharLevelArrayIndex As Integer
        If intLevel = 1 Then
            'level 1 characters should obviously store data in the first index of the array
            intCharLevelArrayIndex = 0
        Else
            intCharLevelArrayIndex = intLevel / 10
        End If

        'ARRAYS---------------------------------------------
        Dim dblArrayBaseHp(9) As Double
        Dim dblArrayBaseAtk(9) As Double
        Dim dblArrayBaseDef(9) As Double
        Dim dblArrayBaseEM(9) As Double

        Dim dblArrayBaseCRate(9) As Double
        Dim dblArrayBaseCDmg(9) As Double
        Dim dblArrayBaseER(9) As Double
        Dim dblArrayBaseHealBonus(9) As Double

        'add more elements here if needed
        Dim dblArrayBaseGeoDmgBonus(9) As Double
        Dim dblArrayBaseHydroDmgBonus(9) As Double

        'MAIN VARS-------------------------------------------
        'base stats
        Dim dblBaseHp As Double = 0
        Dim dblBaseAtk As Double = 0
        Dim dblBaseDef As Double = 0
        Dim dblBaseEM As Double = 0

        'advanced stats
        Dim dblBaseCRate As Double = 0
        Dim dblBaseCDmg As Double = 0
        Dim dblBaseER As Double = 100
        Dim dblBaseHealBonus As Double = 0

        'elemental stats
        'add more elements here if you added a new array above
        Dim dblBaseGeoDmgBonus As Double = 0
        Dim dblBaseHydroDmgBonus As Double = 0
        'CHARACTER DATA-------------------------------------------------------------------------------------
        Select Case strCharName
            Case "Hu Tao" 'besto girl 10/10 ign "he don't lie" 👍👍👍
                'MAIN CHARACTER DATA
                dblArrayBaseHp = {1211, 2070.5, 3141, 5112.3, 6253, 8042, 10089, 11899, 13721, 15552}
                dblArrayBaseAtk = {8, 14.3, 21, 36.8, 43, 55, 69, 81, 94, 106}
                dblArrayBaseDef = {68, 122.54, 177, 274.2, 352, 453, 568, 670, 773, 876}
                'Hu Tao's BaseEM doesn't change when levelling up
                'Hu Tao's BaseCRate doesn't change when levelling up
                dblArrayBaseCDmg = {50, 50, 50, 50, 50, 59.6, 69.2, 69.2, 78.8, 88.4}
                'Hu Tao's BaseER doesn't change when levelling up

                'OUTPUT VARS
                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)

                dblBaseCRate = 5
                dblBaseCDmg = dblArrayBaseCDmg(intCharLevelArrayIndex)

            Case "Albedo"
                'level {1, 10, 20, 30, 40, 50, 60, 70, 80, 90}
                dblArrayBaseHp = {1030, 1850.53, 2671, 4533.4, 5317, 6839, 8579, 10119, 11669, 13226}
                dblArrayBaseAtk = {20, 35.48, 51, 85.9, 101, 130, 163, 192, 222, 251}
                dblArrayBaseDef = {68, 122.45, 177, 300, 352, 453, 568, 670, 773, 876}
                dblArrayBaseGeoDmgBonus = {0, 0, 0, 0, 0, 7.2, 14.4, 14.4, 21.6, 28.8}

                'Output stats
                'Base Statz
                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                'advanced stats
                dblBaseCRate = 5
                dblBaseCDmg = 50
                'elemental stats
                dblBaseGeoDmgBonus = dblArrayBaseGeoDmgBonus(intCharLevelArrayIndex)

            Case "Eula"
                dblArrayBaseHp = {1020, 1850.53, 2671, 4533.4, 5317, 6839, 8579, 10119, 11669, 13226}
                dblArrayBaseAtk = {27, 47.97, 69, 117.6, 138, 177, 222, 262, 302, 342}
                dblArrayBaseDef = {58, 104.98, 152, 257.6, 302, 388, 487, 574, 662, 751}
                dblArrayBaseCDmg = {50, 50, 50, 50, 50, 59.6, 69.2, 69.2, 78.8, 88.4}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = 5
                dblBaseCDmg = dblArrayBaseCDmg(intCharLevelArrayIndex)

            Case "Kaedehara Kazuha"
                dblArrayBaseHp = {1039, 1867, 2695, 4574.89, 5366, 6902, 8659, 10213, 11777, 13348}
                dblArrayBaseAtk = {23, 41.5, 60, 100.56, 119, 153, 192, 227, 262, 297}
                dblArrayBaseDef = {63, 113, 163, 217, 324, 417, 523, 617, 712, 807}
                dblArrayBaseEM = {0, 0, 0, 0, 0, 28.8, 57.6, 57.6, 86.4, 115.2}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = 5
                dblBaseCDmg = 50

            Case "Kamisato Ayaka"
                dblArrayBaseHp = {1001, 1799, 2597, 4341, 5170, 6649, 8341, 9838, 11345, 12858}
                dblArrayBaseAtk = {27, 48, 69, 115.33, 138, 177, 222, 262, 302, 342}
                dblArrayBaseDef = {61, 109.5, 158, 264.89, 315, 405, 509, 600, 692, 784}
                dblArrayBaseCDmg = {50, 50, 50, 50, 50, 59.6, 69.2, 69.2, 78.8, 88.4}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = 5
                dblBaseCDmg = dblArrayBaseCDmg(intCharLevelArrayIndex)

            Case "Kamisato Ayato"
                dblArrayBaseHp = {1068, 1919, 2770, 4630.56, 5514, 7092, 8897, 10494, 12101, 13715}
                dblArrayBaseAtk = {23, 41.5, 60, 100.56, 120, 155, 194, 229, 264, 299}
                dblArrayBaseDef = {60, 107.5, 155, 258.78, 309, 397, 499, 588, 678, 769}
                dblArrayBaseCDmg = {50, 50, 50, 50, 50, 59.6, 69.2, 69.2, 78.8, 88.4}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = 5
                dblBaseCDmg = dblArrayBaseCDmg(intCharLevelArrayIndex)

            Case "Sangonomiya Kokomi"
                dblArrayBaseHp = {1049, 1884.5, 2720, 4547.33, 5416, 6966, 8738, 10306, 11885, 13471}
                dblArrayBaseAtk = {18, 32.5, 47, 79.11, 94, 121, 152, 179, 207, 234}
                dblArrayBaseDef = {51, 92, 133, 222.56, 264, 340, 426, 503, 580, 657}
                dblArrayBaseHydroDmgBonus = {0, 0, 0, 0, 0, 7.2, 14.4, 14.4, 21.6, 28.8}
                'lv60 ascension passive
                dblArrayBaseCRate = {5, 5, 5, 5, 5, 5, -95, -95, -95, -95}
                dblArrayBaseHealBonus = {0, 0, 0, 0, 0, 0, 25, 25, 25, 25}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = dblArrayBaseCRate(intCharLevelArrayIndex)
                dblBaseCDmg = 50
                dblBaseHealBonus = dblArrayBaseHealBonus(intCharLevelArrayIndex)
                dblBaseHydroDmgBonus = dblArrayBaseHydroDmgBonus(intCharLevelArrayIndex)

            Case "Xiao"
                dblArrayBaseHp = {991, 1781.5, 2572, 4300.33, 5120, 6586, 8262, 9744, 11236, 12736}
                dblArrayBaseAtk = {27, 49, 71, 118.44, 140, 181, 227, 267, 308, 349}
                dblArrayBaseDef = {62, 111.5, 161, 273.9, 321, 413, 519, 612, 705, 799}
                dblArrayBaseCRate = {5, 5, 5, 5, 5, 9.8, 14.6, 14.6, 19.4, 24.2}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = dblArrayBaseCRate(intCharLevelArrayIndex)
                dblBaseCDmg = 50

            Case "Yae Miko"
                dblArrayBaseHp = {807, 1451, 2095, 3502.56, 4170, 5364, 6729, 7936, 9151, 10372}
                dblArrayBaseAtk = {26, 47.5, 69, 114.89, 137, 176, 220, 260, 300, 340}
                dblArrayBaseDef = {44, 79.5, 115, 192.44, 229, 294, 369, 435, 502, 569}
                dblArrayBaseCRate = {5, 5, 5, 5, 5, 9.8, 14.6, 14.6, 19.4, 24.2}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = dblArrayBaseCRate(intCharLevelArrayIndex)
                dblBaseCDmg = 50

            Case "Zhongli"
                dblArrayBaseHp = {1144, 2055.5, 2967, 4960.78, 5908, 7599, 9533, 11243, 12965, 14695}
                dblArrayBaseAtk = {20, 35.5, 51, 84.22, 101, 130, 163, 192, 222, 251}
                dblArrayBaseDef = {57, 103, 149, 249.11, 297, 382, 479, 564, 651, 738}
                dblArrayBaseGeoDmgBonus = {0, 0, 0, 0, 0, 7.2, 14.4, 14.4, 21.6, 28.8}

                dblBaseHp = dblArrayBaseHp(intCharLevelArrayIndex)
                dblBaseAtk = dblArrayBaseAtk(intCharLevelArrayIndex)
                dblBaseDef = dblArrayBaseDef(intCharLevelArrayIndex)
                dblBaseCRate = 5
                dblBaseCDmg = 50

        End Select
        Call DisplayBaseStats(dblBaseHp, dblBaseAtk, dblBaseDef)

        'WEAPON SELECTION-------------------------------------------------------------------------------
        Dim intWeapLevelArrayIndex As Integer
        If intWeapLevel = 1 Then
            intWeapLevelArrayIndex = 0
        Else
            intWeapLevelArrayIndex = intWeapLevel / 10
        End If

        Dim intWeapRefineArrayIndex As Integer = 0
        intWeapRefineArrayIndex = intWeapRefine - 1

        'ARRAYS
        Dim dblArrayWeaponAtk(9) As Double
        Dim dblArrayWeaponSubstatVal(9) As Double
        Dim dblArrayWeaponRefinement(4) As Double

        'MAIN VARS
        Dim dblWeaponAtk As Double = 0
        Dim strWeaponSubstatName As String = "Default"
        Dim dblWeaponSubstatVal As Double = 0

        'weapon refinement bonuses - some weapons have it, most don't
        Dim strWeaponRefinementName As String = "Default"
        Dim dblWeaponRefinementStatVal As Double = 0

        Select Case strWeaponName
            Case "Beginner's Protector"
                dblArrayWeaponAtk = {23, 39.47, 56, 86.9, 102, 130, 158, 185, 0, 0}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "No Substat"
                dblWeaponSubstatVal = 0

            Case "Mistsplitter Reforged"
                dblArrayWeaponAtk = {48, 90.48, 133, 217.8, 261, 341, 423, 506, 590, 674}
                dblArrayWeaponSubstatVal = {9.6, 13.29, 17, 21.3, 24.7, 28.6, 32.5, 36.4, 40.2, 44.1}
                dblArrayWeaponRefinement = {12, 15, 18, 21, 24}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "CRIT DMG"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

                strWeaponRefinementName = "Elemental DMG Bonus"
                dblWeaponRefinementStatVal = dblArrayWeaponRefinement(intWeapRefineArrayIndex)
            Case "Staff of Homa"
                dblArrayWeaponAtk = {46, 82, 122, 194, 235, 308, 382, 457, 532, 608}
                dblArrayWeaponSubstatVal = {14.4, 19.6, 25.4, 31.3, 37.1, 42.9, 48.7, 54.5, 60.3, 66.2}
                dblArrayWeaponRefinement = {20, 25, 30, 35, 40}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "CRIT DMG"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

                strWeaponRefinementName = "HP%"
                dblWeaponRefinementStatVal = dblArrayWeaponRefinement(intWeapRefineArrayIndex)

            Case "Amenoma Kageuchi"
                dblArrayWeaponAtk = {41, 69, 99, 155, 184, 238, 293, 347, 401, 454}
                dblArrayWeaponSubstatVal = {12, 16.4, 21.2, 26.1, 30.9, 35.7, 40.6, 45.4, 50.3, 55.1}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "ATK%"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)
            Case "Sacrificial Sword"
                dblArrayWeaponAtk = {41, 69, 99, 155, 184, 238, 293, 347, 401, 454}
                dblArrayWeaponSubstatVal = {13.4, 18.2, 23.6, 28.9, 34.3, 39.7, 45.1, 50.5, 55.9, 61.3}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "Energy Recharge"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

            Case "The Catch"
                dblArrayWeaponAtk = {42, 74, 109, 170, 205, 266, 327, 388, 449, 510}
                dblArrayWeaponSubstatVal = {10, 13.6, 17.7, 21.7, 25.8, 29.8, 33.8, 37.9, 41.9, 45.9}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "Energy Recharge"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

            Case "Song of Broken Pines"
                dblArrayWeaponAtk = {49, 93, 145, 230, 286, 374, 464, 555, 648, 741}
                dblArrayWeaponSubstatVal = {4.5, 6.1, 8, 9.8, 11.6, 13.4, 15.2, 17, 18.9, 20.7}
                dblArrayWeaponRefinement = {16, 20, 24, 28, 32}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "Physical DMG Bonus"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

                strWeaponRefinementName = "ATK%"
                dblWeaponRefinementStatVal = dblArrayWeaponRefinement(intWeapRefineArrayIndex)

            Case "Wolf's Gravestone"
                dblArrayWeaponAtk = {46, 82, 122, 194, 235, 308, 382, 457, 532, 608}
                dblArrayWeaponSubstatVal = {10.8, 14.7, 19.1, 23.4, 27.8, 32.2, 36.5, 40.9, 45.3, 49.6}
                dblArrayWeaponRefinement = {20, 25, 30, 35, 40}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "ATK%"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

                strWeaponRefinementName = "ATK%"
                dblWeaponRefinementStatVal = dblArrayWeaponRefinement(intWeapRefineArrayIndex)

            Case "Everlasting Moonglow"
                dblArrayWeaponAtk = {46, 82, 122, 194, 235, 308, 382, 457, 532, 608}
                dblArrayWeaponSubstatVal = {10.8, 14.7, 19.1, 23.4, 27.8, 32.2, 36.5, 40.9, 45.3, 49.6}
                dblArrayWeaponRefinement = {10, 12.5, 15, 17.5, 20}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "HP%"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

                strWeaponRefinementName = "Healing Bonus"
                dblWeaponRefinementStatVal = dblArrayWeaponRefinement(intWeapRefineArrayIndex)

            Case "Oathsworn Eye"
                dblArrayWeaponAtk = {44, 79, 119, 185, 226, 293, 361, 429, 497, 565}
                dblArrayWeaponSubstatVal = {6, 8.2, 10.6, 13, 15.5, 17.9, 20.3, 22.7, 25.1, 27.6}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "HP%"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

            Case "Luxurious Sea-Lord"
                dblArrayWeaponAtk = {41, 69, 99, 155, 184, 238, 293, 347, 401, 454}
                dblArrayWeaponSubstatVal = {12, 16.4, 21.2, 26.1, 30.9, 35.7, 40.6, 45.4, 50.3, 55.1}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "ATK%"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)

            Case "Primordial Jade-Winged Spear"
                dblArrayWeaponAtk = {48, 87, 133, 212, 261, 341, 423, 506, 590, 674}
                dblArrayWeaponSubstatVal = {4.8, 6.5, 8.5, 10.4, 12.4, 14.3, 16.2, 18.2, 20.1, 22.1}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "CRIT Rate"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)
            Case "Kagura's Verity"
                dblArrayWeaponAtk = {46, 82, 122, 194, 235, 308, 382, 457, 532, 608}
                dblArrayWeaponSubstatVal = {14.4, 19.6, 25.4, 31.3, 37.1, 42.9, 48.7, 54.5, 60.3, 66.2}

                dblWeaponAtk = dblArrayWeaponAtk(intWeapLevelArrayIndex)
                strWeaponSubstatName = "CRIT DMG"
                dblWeaponSubstatVal = dblArrayWeaponSubstatVal(intWeapLevelArrayIndex)
        End Select
        Call DisplayWeaponStats(dblWeaponAtk, dblWeaponSubstatVal, strWeaponSubstatName)
        Call RefinementBonusChecker(dblWeaponRefinementStatVal)
        'adding the base weapon atk to the base char atk
        dblBaseAtk = dblBaseAtk + dblWeaponAtk 'doin' the sneaky

        'TOTAL STATS--------------------------------------------------------------------------------------
        'DECLARATION OF VARIABLES------------------------------------------------------
        'flat stats---------------------------------------------------------------------
        Dim dblTotalHP As Double = 0
        Dim dblTotalAtk As Double = 0
        Dim dblTotalDef As Double = 0
        Dim dblTotalEM As Double = 0

        'percentage based stats--------------------------------------------------------
        Dim dblTotalCRate As Double = 0
        Dim dblTotalCDmg As Double = 0
        Dim dblTotalHealBonus As Double = 0
        Dim dblTotalER As Double = 0
        Dim dblTotalShieldStrength As Double = 0

        Dim dblTotalPyroBonus As Double = 0
        Dim dblTotalPyroRes As Double = 0
        Dim dblTotalHydroBonus As Double = 0
        Dim dblTotalHydroRes As Double = 0
        Dim dblTotalElectroBonus As Double = 0
        Dim dblTotalElectroRes As Double = 0
        Dim dblTotalAnemoBonus As Double = 0
        Dim dblTotalAnemoRes As Double = 0
        Dim dblTotalCryoBonus As Double = 0
        Dim dblTotalCryoRes As Double = 0
        Dim dblTotalGeoBonus As Double = 0
        Dim dblTotalGeoRes As Double = 0
        Dim dblTotalPhysicalBonus As Double = 0
        Dim dblTotalPhysicalRes As Double = 0

        'ADDING THE BASE STATS TO THE 'TOTAL' VARIABLES-------------------------------------------------------
        'hindi talaga '+=' yung nilagay ko dito, huwag pakialaman -Gene
        dblTotalHP = dblBaseHp
        dblTotalAtk = dblBaseAtk
        dblTotalDef = dblBaseDef
        dblTotalEM = dblBaseEM

        dblTotalCRate += dblBaseCRate
        dblTotalCDmg += dblBaseCDmg
        dblTotalER += dblBaseER
        dblTotalHealBonus += dblBaseHealBonus

        'ADD NEW ELEMENTS HERE IF YOU ADDED NEW ONES ON THE BASE STATS
        'like dblTotalPyroBonus, dblTotalAnemoBonus, etc.
        dblTotalGeoBonus += dblBaseGeoDmgBonus
        dblTotalHydroBonus += dblBaseHydroDmgBonus

        'SELECTING THE WEAPON SUBSTAT
        Select Case strWeaponSubstatName
            Case "HP%"
                Dim dblWeaponHPPercentage As Double = 0
                dblWeaponHPPercentage = dblBaseHp * (dblWeaponSubstatVal / 100)
                dblTotalHP += dblWeaponHPPercentage
            Case "ATK%"
                Dim dblWeaponAtkPercentage As Double = 0
                dblWeaponAtkPercentage = dblBaseAtk * (dblWeaponSubstatVal / 100)
                dblTotalAtk += dblWeaponAtkPercentage
            Case "CRIT DMG"
                dblTotalCDmg += dblWeaponSubstatVal
            Case "CRIT Rate"
                dblTotalCRate += dblWeaponSubstatVal
            Case "Energy Recharge"
                dblTotalER += dblWeaponSubstatVal
            Case "Physical DMG Bonus"
                dblTotalPhysicalBonus += dblWeaponSubstatVal
        End Select

        'SELECTING THE WEAPON REFINE BONUS
        Select Case strWeaponRefinementName
            Case "ATK%"
                Dim dblWeaponRefineATKPercentage As Double = 0
                dblWeaponRefineATKPercentage = dblBaseAtk * (dblWeaponRefinementStatVal / 100)
                dblTotalAtk += dblWeaponRefineATKPercentage
            Case "HP%"
                Dim dblWeaponRefineHPPercentage As Double = 0
                dblWeaponRefineHPPercentage = dblBaseHp * (dblWeaponRefineHPPercentage / 100)
                dblTotalHP += dblWeaponRefineHPPercentage
            Case "Healing Bonus"
                dblTotalHealBonus += dblWeaponRefinementStatVal
            Case "Elemental DMG Bonus"
                dblTotalPyroBonus += dblWeaponRefinementStatVal
                dblTotalCryoBonus += dblWeaponRefinementStatVal
                dblTotalAnemoBonus += dblWeaponRefinementStatVal
                dblTotalElectroBonus += dblWeaponRefinementStatVal
                dblTotalHydroBonus += dblWeaponRefinementStatVal
                dblTotalGeoBonus += dblWeaponRefinementStatVal
        End Select


        'MAIN ARTIFACT CALCULATIONS 🌸🪶⌛☕👑-----------------------------------------------------------------------------------
        'ARTIFACT SET COUNTER---------------------------------------------------------------------------
        'Counts if there are any Artifacts coming from the same set

        'setting all counters to 0
        Dim intArchaicPetraCounter As Integer = 0
        Dim intCrimsonWitchCounter As Integer = 0
        Dim intEmblemCounter As Integer = 0
        Dim intGladiatorCounter As Integer = 0
        Dim intShimenawaCounter As Integer = 0
        Dim intWanderersCounter As Integer = 0
        Dim intBlizzardCounter As Integer = 0
        Dim intBloodstainedCounter As Integer = 0

        strArtifactFlower = cmbFlowerSet.Text
        strArtifactPlume = cmbPlumeSet.Text
        strArtifactSands = cmbSandsSet.Text
        strArtifactGoblet = cmbGobletSet.Text
        strArtifactCirclet = cmbCircletSet.Text

        'storing the artifact names in an array makes the case selection loopable, which shortens the code
        Dim strArrayFiveArtifacts(4) As String
        strArrayFiveArtifacts(0) = strArtifactFlower
        strArrayFiveArtifacts(1) = strArtifactPlume
        strArrayFiveArtifacts(2) = strArtifactSands
        strArrayFiveArtifacts(3) = strArtifactGoblet
        strArrayFiveArtifacts(4) = strArtifactCirclet

        For intArtifactIndex As Integer = 0 To 4 Step 1
            Select Case strArrayFiveArtifacts(intArtifactIndex)
                Case "Archaic Petra"
                    intArchaicPetraCounter += 1
                Case "Crimson Witch of Flames"
                    intCrimsonWitchCounter += 1
                Case "Emblem of Severed Fates"
                    intEmblemCounter += 1
                Case "Gladiator's Finale"
                    intGladiatorCounter += 1
                Case "Shimenawa's Reminiscence"
                    intShimenawaCounter += 1
                Case "Wanderer's Troupe"
                    intWanderersCounter += 1
                Case "Blizzard Strayer"
                    intBlizzardCounter += 1
                Case "Bloodstained Chivalry"
                    intBloodstainedCounter += 1
            End Select
        Next intArtifactIndex

        'Debugging
        Debug.WriteLine("[ArtifactCalculate]:Archaic Petra Counter: " + intArchaicPetraCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Crimson Witch Counter: " + intCrimsonWitchCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Emblem Counter: " + intEmblemCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Gladiator Counter: " + intGladiatorCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Shimenawa Counter: " + intShimenawaCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Wanderers Counter: " + intWanderersCounter.ToString)
        Debug.WriteLine("[ArtifactCalculate]:Blizzard Strayer Counter: " + intBlizzardCounter.ToString)

        'ARTIFACT SET BONUS CALCULATOR--------------------------------------------------------------------------
        lstArtifactBonuses.Items.Clear() 'clear the list box first; refresh the list box
        If intArchaicPetraCounter >= 2 Then
            dblTotalGeoBonus += 15.0
            lstArtifactBonuses.Items.Add("Archaic Petra (2): +15% Geo DMG Bonus")
        End If
        If intCrimsonWitchCounter >= 2 Then
            dblTotalPyroBonus += 15.0
            lstArtifactBonuses.Items.Add("Crimson Witch of Flames (2): +15% Pyro Bonus")
        End If
        If intEmblemCounter >= 2 Then
            dblTotalER += 20.0
            lstArtifactBonuses.Items.Add("Emblem of Severed Fates (2): +20% Energy Recharge")
        End If
        If intGladiatorCounter >= 2 Then
            Dim dblGladAtkPercentage As Double = 0
            dblGladAtkPercentage = dblBaseAtk * 0.18
            dblTotalAtk += dblGladAtkPercentage
            lstArtifactBonuses.Items.Add("Gladiator's Finale (2): +18% ATK")
        End If
        If intWanderersCounter >= 2 Then
            dblTotalEM += 80
            lstArtifactBonuses.Items.Add("Wanderer's Troupe (2): +80 Elemental Mastery")
        End If
        If intShimenawaCounter >= 2 Then
            Dim dblSMNWAtkPercentage As Double = 0
            dblSMNWAtkPercentage = dblBaseAtk * 0.18
            dblTotalAtk += dblSMNWAtkPercentage
            lstArtifactBonuses.Items.Add("Shimenawa's Reminiscence (2): +18% ATK")
        End If
        If intBlizzardCounter >= 2 Then
            dblTotalCryoBonus += 15
            lstArtifactBonuses.Items.Add("Blizzard Strayer (2): +15% Cryo DMG Bonus")
        End If
        If intBloodstainedCounter >= 2 Then
            dblTotalPhysicalBonus += 25
            lstArtifactBonuses.Items.Add("Bloodstained Chivalry (2): +25% Physical DMG Bonus")
        End If

        'ARTIFACT SUBSTATS CALCULATOR-----------------------------------------------------------------------
        'Taking the inputted stats on the 'Edit Artifacts' panel and adding it to the TOTAL stats-----------
        'FLOWER 🌸 ----------------------------------------------------------------------------------------
        Dim dblFlowerHP As Double = 0
        Dim dblFlowerSubstat1Value As Double = 0
        Dim dblFlowerSubstat2Value As Double = 0
        Dim dblFlowerSubstat3Value As Double = 0
        Dim dblFlowerSubstat4Value As Double = 0
        Double.TryParse(txtFlowerHP.Text, dblFlowerHP)
        Double.TryParse(txtFlowerSubstat1Value.Text, dblFlowerSubstat1Value)
        Double.TryParse(txtFlowerSubstat2Value.Text, dblFlowerSubstat2Value)
        Double.TryParse(txtFlowerSubstat3Value.Text, dblFlowerSubstat3Value)
        Double.TryParse(txtFlowerSubstat4Value.Text, dblFlowerSubstat4Value)

        Dim strFlowerSubstatNames(3) As String
        strFlowerSubstatNames(0) = cmbFlowerSubstat1Name.Text.ToLower
        strFlowerSubstatNames(1) = cmbFlowerSubstat2Name.Text.ToLower
        strFlowerSubstatNames(2) = cmbFlowerSubstat3Name.Text.ToLower
        strFlowerSubstatNames(3) = cmbFlowerSubstat4Name.Text.ToLower

        Dim dblFlowerSubstatValues(3) As Double
        dblFlowerSubstatValues(0) = dblFlowerSubstat1Value
        dblFlowerSubstatValues(1) = dblFlowerSubstat2Value
        dblFlowerSubstatValues(2) = dblFlowerSubstat3Value
        dblFlowerSubstatValues(3) = dblFlowerSubstat4Value

        'Debugging---------------------
        Debug.WriteLine(vbCrLf + "[ArtifactFlower]flowermainstat(hp) is " + dblFlowerHP.ToString)
        For intDebugCounter As Integer = 0 To 3 Step 1
            Debug.WriteLine("[ArtifactPlume]flowersubstat" + (intDebugCounter + 1).ToString + " is " +
                            strFlowerSubstatNames(intDebugCounter) + " with a value of " +
                            dblFlowerSubstatValues(intDebugCounter).ToString)
        Next intDebugCounter
        '-----------------------------

        dblTotalHP = dblTotalHP + dblFlowerHP 'add the flowers main stat to the total stats
        For intSubstatCounter As Integer = 0 To 3 Step 1
            Select Case strFlowerSubstatNames(intSubstatCounter)
                Case "hp%"
                    Dim dblFlowerHPPercentage As Double = 0
                    dblFlowerHPPercentage = dblBaseHp * (dblFlowerSubstatValues(intSubstatCounter) / 100)
                    dblTotalHP = dblTotalHP + dblFlowerHPPercentage
                Case "atk"
                    dblTotalAtk = dblTotalAtk + dblFlowerSubstatValues(intSubstatCounter)
                Case "atk%"
                    Dim dblFlowerATKPercentage As Double = 0
                    dblFlowerATKPercentage = dblBaseAtk * (dblFlowerSubstatValues(intSubstatCounter) / 100)
                    dblTotalAtk = dblTotalAtk + dblFlowerATKPercentage
                Case "def"
                    dblTotalDef = dblTotalDef + dblFlowerSubstatValues(intSubstatCounter)
                Case "def%"
                    Dim dblFlowerDEFPercentage As Double = 0
                    dblFlowerDEFPercentage = dblBaseDef * (dblFlowerSubstatValues(intSubstatCounter) / 100)
                    dblTotalDef = dblTotalDef + dblFlowerDEFPercentage
                Case "crit rate"
                    dblTotalCRate = dblTotalCRate + dblFlowerSubstatValues(intSubstatCounter)
                Case "crit dmg"
                    dblTotalCDmg = dblTotalCDmg + dblFlowerSubstatValues(intSubstatCounter)
                Case "elemental mastery"
                    dblTotalEM = dblTotalEM + dblFlowerSubstatValues(intSubstatCounter)
                Case "energy recharge"
                    dblTotalER = dblTotalER + dblFlowerSubstatValues(intSubstatCounter)
            End Select
        Next intSubstatCounter

        'PLUME 🪶 ----------------------------------------------------------------------------------------
        Dim dblPlumeAtk As Double = 0
        Dim dblPlumeSubstat1Value As Double = 0
        Dim dblPlumeSubstat2Value As Double = 0
        Dim dblPlumeSubstat3Value As Double = 0
        Dim dblPlumeSubstat4Value As Double = 0
        Double.TryParse(txtPlumeATK.Text, dblPlumeAtk)
        Double.TryParse(txtPlumeSubstat1Value.Text, dblPlumeSubstat1Value)
        Double.TryParse(txtPlumeSubstat2Value.Text, dblPlumeSubstat2Value)
        Double.TryParse(txtPlumeSubstat3Value.Text, dblPlumeSubstat3Value)
        Double.TryParse(txtPlumeSubstat4Value.Text, dblPlumeSubstat4Value)

        Dim strPlumeSubstatNames(3) As String
        strPlumeSubstatNames(0) = cmbPlumeSubstat1Name.Text.ToLower
        strPlumeSubstatNames(1) = cmbPlumeSubstat2Name.Text.ToLower
        strPlumeSubstatNames(2) = cmbPlumeSubstat3Name.Text.ToLower
        strPlumeSubstatNames(3) = cmbPlumeSubstat4Name.Text.ToLower

        Dim dblPlumeSubstatValues(3) As Double
        dblPlumeSubstatValues(0) = dblPlumeSubstat1Value
        dblPlumeSubstatValues(1) = dblPlumeSubstat2Value
        dblPlumeSubstatValues(2) = dblPlumeSubstat3Value
        dblPlumeSubstatValues(3) = dblPlumeSubstat4Value

        'Debugging
        Debug.WriteLine(vbCrLf + "[ArtifactPlume]plumemainstat(atk) is " + dblPlumeAtk.ToString)
        For intDebugCounter As Integer = 0 To 3 Step 1
            Debug.WriteLine("[ArtifactPlume]plumesubstat" + (intDebugCounter + 1).ToString +
                            " is " + strPlumeSubstatNames(intDebugCounter) + " with a value of " +
                            dblPlumeSubstatValues(intDebugCounter).ToString)
        Next intDebugCounter

        dblTotalAtk = dblTotalAtk + dblPlumeAtk
        For intSubstatCounter As Integer = 0 To 3 Step 1
            Select Case strPlumeSubstatNames(intSubstatCounter)
                Case "hp"
                    dblTotalHP = dblTotalHP + dblPlumeSubstatValues(intSubstatCounter)
                Case "hp%"
                    Dim dblPlumeHPPercentage As Double = 0
                    dblPlumeHPPercentage = dblBaseHp * (dblPlumeSubstatValues(intSubstatCounter) / 100)
                    dblTotalHP = dblTotalHP + dblPlumeHPPercentage
                Case "atk%"
                    Dim dblPlumeATKPercentage As Double = 0
                    dblPlumeATKPercentage = dblBaseAtk * (dblPlumeSubstatValues(intSubstatCounter) / 100)
                    dblTotalAtk = dblTotalAtk + dblPlumeATKPercentage
                Case "def"
                    dblTotalDef = dblTotalDef + dblPlumeSubstatValues(intSubstatCounter)
                Case "def%"
                    Dim dblPlumeDEFPercentage As Double = 0
                    dblPlumeDEFPercentage = dblBaseDef * (dblPlumeSubstatValues(intSubstatCounter) / 100)
                    dblTotalDef = dblTotalDef + dblPlumeDEFPercentage
                Case "crit rate"
                    dblTotalCRate = dblTotalCRate + dblPlumeSubstatValues(intSubstatCounter)
                Case "crit dmg"
                    dblTotalCDmg = dblTotalCDmg + dblPlumeSubstatValues(intSubstatCounter)
                Case "elemental mastery"
                    dblTotalEM = dblTotalEM + dblPlumeSubstatValues(intSubstatCounter)
                Case "energy recharge"
                    dblTotalER = dblTotalER + dblPlumeSubstatValues(intSubstatCounter)
            End Select
        Next intSubstatCounter

        'SANDS ⌛ ----------------------------------------------------------------------------------------
        Dim dblSandsMainStatValue As Double = 0
        Dim dblSandsSubstat1Value As Double = 0
        Dim dblSandsSubstat2Value As Double = 0
        Dim dblSandsSubstat3Value As Double = 0
        Dim dblSandsSubstat4Value As Double = 0
        Double.TryParse(txtSandsMainStatValue.Text, dblSandsMainStatValue)
        Double.TryParse(txtSandsSubstat1Value.Text, dblSandsSubstat1Value)
        Double.TryParse(txtSandsSubstat2Value.Text, dblSandsSubstat2Value)
        Double.TryParse(txtSandsSubstat3Value.Text, dblSandsSubstat3Value)
        Double.TryParse(txtSandsSubstat4Value.Text, dblSandsSubstat4Value)

        Dim strSandsStatNames(4) As String
        strSandsStatNames(0) = cmbSandsMainStatName.Text.ToLower 'MAIN STAT
        strSandsStatNames(1) = cmbSandsSubstat1Name.Text.ToLower
        strSandsStatNames(2) = cmbSandsSubstat2Name.Text.ToLower
        strSandsStatNames(3) = cmbSandsSubstat3Name.Text.ToLower
        strSandsStatNames(4) = cmbSandsSubstat4Name.Text.ToLower

        Dim dblSandsStatValues(4) As String
        dblSandsStatValues(0) = dblSandsMainStatValue 'MAIN STAT
        dblSandsStatValues(1) = dblSandsSubstat1Value
        dblSandsStatValues(2) = dblSandsSubstat2Value
        dblSandsStatValues(3) = dblSandsSubstat3Value
        dblSandsStatValues(4) = dblSandsSubstat4Value

        'debugging
        Debug.WriteLine(vbCrLf + "[ArtifactSands]sandsmainstat(" + strSandsStatNames(0) + ") is " +
                        dblSandsStatValues(0).ToString)
        For intDebugCounter As Integer = 0 To 4 Step 1
            Debug.WriteLine("[ArtifactSands]sandssubstat" + (intDebugCounter + 1).ToString + " is " +
                            strSandsStatNames(intDebugCounter) + " with a value of " +
                            dblSandsStatValues(intDebugCounter).ToString)
        Next intDebugCounter

        For intSubstatCounter As Integer = 0 To 4 Step 1
            Select Case strSandsStatNames(intSubstatCounter)
                Case "hp"
                    dblTotalHP = dblTotalHP + dblSandsStatValues(intSubstatCounter)
                Case "hp%"
                    Dim dblSandsHPPercentage As Double = 0
                    dblSandsHPPercentage = dblBaseHp * (dblSandsStatValues(intSubstatCounter) / 100)
                    dblTotalHP = dblTotalHP + dblSandsHPPercentage
                Case "atk"
                    dblTotalAtk = dblTotalAtk + dblSandsStatValues(intSubstatCounter)
                Case "atk%"
                    Dim dblSandsATKPercentage As Double = 0
                    dblSandsATKPercentage = dblBaseAtk * (dblSandsStatValues(intSubstatCounter) / 100)
                    dblTotalAtk = dblTotalAtk + dblSandsATKPercentage
                Case "def"
                    dblTotalDef = dblTotalDef + dblSandsStatValues(intSubstatCounter)
                Case "def%"
                    Dim dblSandsDEFPercentage As Double = 0
                    dblSandsDEFPercentage = dblBaseDef * (dblSandsStatValues(intSubstatCounter) / 100)
                    dblTotalDef = dblTotalDef + dblSandsDEFPercentage
                Case "crit rate"
                    dblTotalCRate = dblTotalCRate + dblSandsStatValues(intSubstatCounter)
                Case "crit dmg"
                    dblTotalCDmg = dblTotalCDmg + dblSandsStatValues(intSubstatCounter)
                Case "elemental mastery"
                    dblTotalEM = dblTotalEM + dblSandsStatValues(intSubstatCounter)
                Case "energy recharge"
                    dblTotalER = dblTotalER + dblSandsStatValues(intSubstatCounter)
            End Select
        Next intSubstatCounter
        'GOBLET ☕ -----------------------------------------------------------------------------------------
        Dim dblGobletMainStatValue As Double = 0
        Dim dblGobletSubstat1Value As Double = 0
        Dim dblGobletSubstat2Value As Double = 0
        Dim dblGobletSubstat3Value As Double = 0
        Dim dblGobletSubstat4Value As Double = 0
        Double.TryParse(txtGobletMainStatValue.Text, dblGobletMainStatValue)
        Double.TryParse(txtGobletSubstat1Value.Text, dblGobletSubstat1Value)
        Double.TryParse(txtGobletSubstat2Value.Text, dblGobletSubstat2Value)
        Double.TryParse(txtGobletSubstat3Value.Text, dblGobletSubstat3Value)
        Double.TryParse(txtGobletSubstat4Value.Text, dblGobletSubstat4Value)

        Dim strGobletStatNames(4) As String
        strGobletStatNames(0) = cmbGobletMainStatName.Text.ToLower
        strGobletStatNames(1) = cmbGobletSubstat1Name.Text.ToLower
        strGobletStatNames(2) = cmbGobletSubstat2Name.Text.ToLower
        strGobletStatNames(3) = cmbGobletSubstat3Name.Text.ToLower
        strGobletStatNames(4) = cmbGobletSubstat4Name.Text.ToLower

        Dim dblGobletStatValues(4) As String
        dblGobletStatValues(0) = dblGobletMainStatValue
        dblGobletStatValues(1) = dblGobletSubstat1Value
        dblGobletStatValues(2) = dblGobletSubstat2Value
        dblGobletStatValues(3) = dblGobletSubstat3Value
        dblGobletStatValues(4) = dblGobletSubstat4Value

        'debugging
        Debug.WriteLine(vbCrLf + "[ArtifactGoblet]gobletmainstat(" + strGobletStatNames(0) + ") is " +
                        dblGobletStatValues(0).ToString)
        For intDebugCounter As Integer = 0 To 4 Step 1
            Debug.WriteLine("[ArtifactGoblet]gobletssubstat" + (intDebugCounter + 1).ToString + " is " +
                            strGobletStatNames(intDebugCounter) + " with a value of " +
                            dblGobletStatValues(intDebugCounter).ToString)
        Next intDebugCounter

        For intSubstatCounter As Integer = 0 To 4 Step 1
            Select Case strGobletStatNames(intSubstatCounter)
                Case "hp"
                    dblTotalHP = dblTotalHP + dblGobletStatValues(intSubstatCounter)
                Case "hp%"
                    Dim dblGobletHPPercentage As Double = 0
                    dblGobletHPPercentage = dblBaseHp * (dblGobletStatValues(intSubstatCounter) / 100)
                    dblTotalHP = dblTotalHP + dblGobletHPPercentage
                Case "atk"
                    dblTotalAtk = dblTotalAtk + dblGobletStatValues(intSubstatCounter)
                Case "atk%"
                    Dim dblGobletATKPercentage As Double = 0
                    dblGobletATKPercentage = dblBaseAtk * (dblGobletStatValues(intSubstatCounter) / 100)
                    dblTotalAtk = dblTotalAtk + dblGobletATKPercentage
                Case "def"
                    dblTotalDef = dblTotalDef + dblGobletStatValues(intSubstatCounter)
                Case "def%"
                    Dim dblGobletDEFPercentage As Double = 0
                    dblGobletDEFPercentage = dblBaseDef * (dblGobletStatValues(intSubstatCounter) / 100)
                    dblTotalDef = dblTotalDef + dblGobletDEFPercentage
                Case "crit rate"
                    dblTotalCRate = dblTotalCRate + dblGobletStatValues(intSubstatCounter)
                Case "crit dmg"
                    dblTotalCDmg = dblTotalCDmg + dblGobletStatValues(intSubstatCounter)
                Case "elemental mastery"
                    dblTotalEM = dblTotalEM + dblGobletStatValues(intSubstatCounter)
                Case "energy recharge"
                    dblTotalER = dblTotalER + dblGobletStatValues(intSubstatCounter)
                Case "physical dmg bonus"
                    dblTotalPhysicalBonus = dblTotalPhysicalBonus + dblGobletStatValues(intSubstatCounter)
                Case "pyro dmg bonus"
                    dblTotalPyroBonus = dblTotalPyroBonus + dblGobletStatValues(intSubstatCounter)
                Case "electro dmg bonus"
                    dblTotalElectroBonus = dblTotalElectroBonus + dblGobletStatValues(intSubstatCounter)
                Case "cryo dmg bonus"
                    dblTotalCryoBonus = dblTotalCryoBonus + dblGobletStatValues(intSubstatCounter)
                Case "hydro dmg bonus"
                    dblTotalHydroBonus = dblTotalHydroBonus + dblGobletStatValues(intSubstatCounter)
                Case "anemo dmg bonus"
                    dblTotalAnemoBonus = dblTotalAnemoBonus + dblGobletStatValues(intSubstatCounter)
                Case "geo dmg bonus"
                    dblTotalGeoBonus = dblTotalGeoBonus + dblGobletStatValues(intSubstatCounter)
            End Select
        Next intSubstatCounter
        'CIRCLET  👑 ---------------------------------------------------------------------------------------
        Dim dblCircletMainStatValue As Double = 0
        Dim dblCircletSubstat1Value As Double = 0
        Dim dblCircletSubstat2Value As Double = 0
        Dim dblCircletSubstat3Value As Double = 0
        Dim dblCircletSubstat4Value As Double = 0
        Double.TryParse(txtCircletMainStatValue.Text, dblCircletMainStatValue)
        Double.TryParse(txtCircletSubstat1Value.Text, dblCircletSubstat1Value)
        Double.TryParse(txtCircletSubstat2Value.Text, dblCircletSubstat2Value)
        Double.TryParse(txtCircletSubstat3Value.Text, dblCircletSubstat3Value)
        Double.TryParse(txtCircletSubstat4Value.Text, dblCircletSubstat4Value)

        Dim strCircletStatNames(4) As String
        strCircletStatNames(0) = cmbCircletMainStatName.Text.ToLower
        strCircletStatNames(1) = cmbCircletSubstat1Name.Text.ToLower
        strCircletStatNames(2) = cmbCircletSubstat2Name.Text.ToLower
        strCircletStatNames(3) = cmbCircletSubstat3Name.Text.ToLower
        strCircletStatNames(4) = cmbCircletSubstat4Name.Text.ToLower

        Dim dblCircletStatValues(4) As Double
        dblCircletStatValues(0) = dblCircletMainStatValue
        dblCircletStatValues(1) = dblCircletSubstat1Value
        dblCircletStatValues(2) = dblCircletSubstat2Value
        dblCircletStatValues(3) = dblCircletSubstat3Value
        dblCircletStatValues(4) = dblCircletSubstat4Value

        'debugging
        Debug.WriteLine(vbCrLf + "[ArtifactCirclet]circletmainstat(" + strCircletStatNames(0) + ") is " +
                        dblCircletStatValues(0).ToString)
        For intDebugCounter As Integer = 0 To 4 Step 1
            Debug.WriteLine("[ArtifactCirclet]circletsubstat" + (intDebugCounter + 1).ToString + " is " +
                            strCircletStatNames(intDebugCounter) + " with a value of " +
                            dblCircletStatValues(intDebugCounter).ToString)
        Next intDebugCounter

        For intSubstatCounter As Integer = 0 To 4 Step 1
            Select Case strCircletStatNames(intSubstatCounter)
                Case "hp"
                    dblTotalHP = dblTotalHP + dblCircletStatValues(intSubstatCounter)
                Case "hp%"
                    Dim dblCircletHPPercentage As Double = 0
                    dblCircletHPPercentage = dblBaseHp * (dblCircletStatValues(intSubstatCounter) / 100)
                    dblTotalHP = dblTotalHP + dblCircletHPPercentage
                Case "atk"
                    dblTotalAtk = dblTotalAtk + dblCircletStatValues(intSubstatCounter)
                Case "atk%"
                    Dim dblCircletATKPercentage As Double = 0
                    dblCircletATKPercentage = dblBaseAtk * (dblCircletStatValues(intSubstatCounter) / 100)
                    dblTotalAtk = dblTotalAtk + dblCircletATKPercentage
                Case "def"
                    dblTotalDef = dblTotalDef + dblCircletStatValues(intSubstatCounter)
                Case "def%"
                    Dim dblCircletDEFPercentage As Double = 0
                    dblCircletDEFPercentage = dblBaseDef * (dblCircletStatValues(intSubstatCounter) / 100)
                    dblTotalDef = dblTotalDef + dblCircletDEFPercentage
                    Debug.WriteLine("[CircletStats]Circlet DEF Percentage is " + dblCircletDEFPercentage.ToString + vbCrLf +
                                    "[CircletStats]Total DEF Percentage is now " + dblTotalDef.ToString)
                Case "crit rate"
                    dblTotalCRate = dblTotalCRate + dblCircletStatValues(intSubstatCounter)
                Case "crit dmg"
                    dblTotalCDmg = dblTotalCDmg + dblCircletStatValues(intSubstatCounter)
                Case "elemental mastery"
                    dblTotalEM = dblTotalEM + dblCircletStatValues(intSubstatCounter)
                Case "energy recharge"
                    dblTotalER = dblTotalER + dblCircletStatValues(intSubstatCounter)
                Case "healing bonus"
                    dblTotalHealBonus = dblTotalHealBonus + dblCircletStatValues(intSubstatCounter)
            End Select
        Next intSubstatCounter

        'DISPLAY ALL THE STATS IN TOTAL---------------------------------------------
        'base stats
        lblMaxHP.Text = Int(dblTotalHP.ToString)
        lblAtk.Text = Int(dblTotalAtk.ToString)
        lblDef.Text = Int(dblTotalDef.ToString)
        lblEm.Text = Int(dblTotalEM.ToString)
        'converting them to int removes their decimals

        'advanced stats
        lblCritRate.Text = dblTotalCRate.ToString("N1") + "%"
        lblCritDmg.Text = dblTotalCDmg.ToString("N1") + "%"
        lblEnergyRecharge.Text = dblTotalER.ToString("N1") + "%"
        lblHealBonus.Text = dblTotalHealBonus.ToString("N1") + "%"
        lblIncomingHeal.Text = "0%"
        lblShieldStrength.Text = "0%"
        'elemental stats
        lblPyroDmgBonus.Text = dblTotalPyroBonus.ToString("N1") + "%"
        lblPyroRes.Text = dblTotalPyroRes.ToString("N1") + "%"
        lblHydroDmgBonus.Text = dblTotalHydroBonus.ToString("N1") + "%"
        lblHydroRes.Text = dblTotalHydroRes.ToString("N1") + "%"
        lblElectroDmgBonus.Text = dblTotalElectroBonus.ToString("N1") + "%"
        lblElectroRes.Text = dblTotalElectroRes.ToString("N1") + "%"
        lblAnemoDmgBonus.Text = dblTotalAnemoBonus.ToString("N1") + "%"
        lblAnemoRes.Text = dblTotalAnemoRes.ToString("N1") + "%"
        lblCryoDmgBonus.Text = dblTotalCryoBonus.ToString("N1") + "%"
        lblCryoRes.Text = dblTotalCryoRes.ToString("N1") + "%"
        lblGeoDmgBonus.Text = dblTotalGeoBonus.ToString("N1") + "%"
        lblGeoRes.Text = dblTotalGeoRes.ToString("N1") + "%"
        lblPhysicalDmgBonus.Text = dblTotalPhysicalBonus.ToString("N1") + "%"
        lblPhysicalRes.Text = dblTotalPhysicalRes.ToString("N1") + "%"
    End Sub

    'This Sub will display the base stats on the Character Stats Group Box
    Private Sub DisplayBaseStats(ByVal dblBaseHP As Double, ByVal dblBaseAtk As Double, ByVal dblBaseDef As Double)
        lblCSMaxHP.Text = Int(dblBaseHP).ToString
        lblCSATK.Text = Int(dblBaseAtk).ToString
        lblCSDEF.Text = Int(dblBaseDef).ToString
    End Sub
    'This Sub will display the base stats on the Weapon Stats Group Box
    Private Sub DisplayWeaponStats(ByVal dblWeapAtk As Double, ByVal dblWeapSubstatValue As Double, ByVal strWeapSubstatName As String)
        lblWeaponAtkValue.Text = Int(dblWeapAtk).ToString

        If dblWeapSubstatValue = 0 Then
            lblWeaponSubStatValue.Text = String.Empty
        Else
            lblWeaponSubStatValue.Text = dblWeapSubstatValue.ToString + "%"
        End If

        lblWeaponSubStatName.Text = strWeapSubstatName
    End Sub

    Private Sub lstPresetList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPresetList.SelectedIndexChanged
        Dim strSelectedPresetBuild As String = lstPresetList.Text
        Select Case strSelectedPresetBuild
            Case "Hu Tao"
                lstCharacters.SelectedItem = "Hu Tao"
                lstLevel.SelectedItem = "90"

                lstWeapons.SelectedItem = "Primordial Jade-Winged Spear"
                lstWeaponLevel.SelectedItem = "90"
                lstWeaponRefinement.SelectedItem = "1"

                cmbFlowerSet.Text = "Gladiator's Finale"
                cmbFlowerSubstat1Name.Text = "CRIT DMG"
                cmbFlowerSubstat2Name.Text = "HP%"
                cmbFlowerSubstat3Name.Text = "CRIT Rate"
                cmbFlowerSubstat4Name.Text = "Energy Recharge"
                txtFlowerHP.Text = "4780"
                txtFlowerSubstat1Value.Text = "21.0"
                txtFlowerSubstat2Value.Text = "9.9"
                txtFlowerSubstat3Value.Text = "6.6"
                txtFlowerSubstat4Value.Text = "11.7"

                cmbPlumeSet.Text = "Shimenawa's Reminiscence"
                cmbPlumeSubstat1Name.Text = "HP%"
                cmbPlumeSubstat2Name.Text = "HP"
                cmbPlumeSubstat3Name.Text = "ATK%"
                cmbPlumeSubstat4Name.Text = "Elemental Mastery"
                txtPlumeATK.Text = "311"
                txtPlumeSubstat1Value.Text = "19.8"
                txtPlumeSubstat2Value.Text = "299"
                txtPlumeSubstat3Value.Text = "8.7"
                txtPlumeSubstat4Value.Text = "23"

                cmbSandsSet.Text = "Shimenawa's Reminiscence"
                cmbSandsMainStatName.Text = "HP%"
                cmbSandsSubstat1Name.Text = "CRIT DMG"
                cmbSandsSubstat2Name.Text = "ATK"
                cmbSandsSubstat3Name.Text = "CRIT Rate"
                cmbSandsSubstat4Name.Text = "Energy Recharge"
                txtSandsMainStatValue.Text = "46.6"
                txtSandsSubstat1Value.Text = "25.6"
                txtSandsSubstat2Value.Text = "37"
                txtSandsSubstat3Value.Text = "7.4"
                txtSandsSubstat4Value.Text = "5.8"

                cmbGobletSet.Text = "Archaic Petra"
                cmbGobletMainStatName.Text = "Pyro DMG Bonus"
                cmbGobletSubstat1Name.Text = "ATK%"
                cmbGobletSubstat2Name.Text = "CRIT Rate"
                cmbGobletSubstat3Name.Text = "Elemental Mastery"
                cmbGobletSubstat4Name.Text = "Energy Recharge"
                txtGobletMainStatValue.Text = "46.6"
                txtGobletSubstat1Value.Text = "10.5"
                txtGobletSubstat2Value.Text = "9.3"
                txtGobletSubstat3Value.Text = "35"
                txtGobletSubstat4Value.Text = "4.5"

                cmbCircletSet.Text = "Gladiator's Finale"
                cmbCircletMainStatName.Text = "CRIT DMG"
                cmbCircletSubstat1Name.Text = "HP"
                cmbCircletSubstat2Name.Text = "DEF"
                cmbCircletSubstat3Name.Text = "ATK"
                cmbCircletSubstat4Name.Text = "ATK%"
                txtCircletMainStatValue.Text = "62.2"
                txtCircletSubstat1Value.Text = "747"
                txtCircletSubstat2Value.Text = "44"
                txtCircletSubstat3Value.Text = "33"
                txtCircletSubstat4Value.Text = "10.5"
            Case "Yae Miko"
                lstCharacters.SelectedItem = "Yae Miko"
                lstLevel.SelectedItem = "90"

                lstWeapons.SelectedItem = "Kagura's Verity"
                lstWeaponLevel.SelectedItem = "90"
                lstWeaponRefinement.SelectedItem = "1"

                cmbFlowerSet.Text = "Shimenawa's Reminiscence"
                cmbFlowerSubstat1Name.Text = "CRIT Rate"
                cmbFlowerSubstat2Name.Text = "CRIT DMG"
                cmbFlowerSubstat3Name.Text = "Elemental Mastery"
                cmbFlowerSubstat4Name.Text = "ATK%"
                txtFlowerHP.Text = "4780"
                txtFlowerSubstat1Value.Text = "12.8"
                txtFlowerSubstat2Value.Text = "5.4"
                txtFlowerSubstat3Value.Text = "65"
                txtFlowerSubstat4Value.Text = "4.7"

                cmbPlumeSet.Text = "Shimenawa's Reminiscence"
                cmbPlumeSubstat1Name.Text = "DEF"
                cmbPlumeSubstat2Name.Text = "CRIT DMG"
                cmbPlumeSubstat3Name.Text = "Energy Recharge"
                cmbPlumeSubstat4Name.Text = "CRIT Rate"
                txtPlumeATK.Text = "311"
                txtPlumeSubstat1Value.Text = "42"
                txtPlumeSubstat2Value.Text = "21.8"
                txtPlumeSubstat3Value.Text = "4.5"
                txtPlumeSubstat4Value.Text = "10.1"

                cmbSandsSet.Text = "Gladiator's Finale"
                cmbSandsMainStatName.Text = "ATK%"
                cmbSandsSubstat1Name.Text = "CRIT DMG"
                cmbSandsSubstat2Name.Text = "Energy Recharge"
                cmbSandsSubstat3Name.Text = "HP"
                cmbSandsSubstat4Name.Text = "CRIT Rate"
                txtSandsMainStatValue.Text = "46.6"
                txtSandsSubstat1Value.Text = "18.7"
                txtSandsSubstat2Value.Text = "10.4"
                txtSandsSubstat3Value.Text = "478"
                txtSandsSubstat4Value.Text = "3.5"

                cmbGobletSet.Text = "Bloodstained Chivalry"
                cmbGobletMainStatName.Text = "Electro DMG Bonus"
                cmbGobletSubstat1Name.Text = "ATK"
                cmbGobletSubstat2Name.Text = "Energy Recharge"
                cmbGobletSubstat3Name.Text = "ATK%"
                cmbGobletSubstat4Name.Text = "HP"
                txtGobletMainStatValue.Text = "46.6"
                txtGobletSubstat1Value.Text = "18"
                txtGobletSubstat2Value.Text = "9.7"
                txtGobletSubstat3Value.Text = "16.3"
                txtGobletSubstat4Value.Text = "418"

                cmbCircletSet.Text = "Gladiator's Finale"
                cmbCircletMainStatName.Text = "CRIT DMG"
                cmbCircletSubstat1Name.Text = "ATK%"
                cmbCircletSubstat2Name.Text = "Energy Recharge"
                cmbCircletSubstat3Name.Text = "ATK"
                cmbCircletSubstat4Name.Text = "Elemental Mastery"
                txtCircletMainStatValue.Text = "62.2"
                txtCircletSubstat1Value.Text = "14.6"
                txtCircletSubstat2Value.Text = "16.8"
                txtCircletSubstat3Value.Text = "19"
                txtCircletSubstat4Value.Text = "21"
            Case "Kamisato Ayaka"
                lstCharacters.SelectedItem = "Kamisato Ayaka"
                lstLevel.SelectedItem = "90"

                lstWeapons.SelectedItem = "Mistsplitter Reforged"
                lstWeaponLevel.SelectedItem = "90"
                lstWeaponRefinement.SelectedItem = "1"

                cmbFlowerSet.Text = "Blizzard Strayer"
                cmbFlowerSubstat1Name.Text = "ATK%"
                cmbFlowerSubstat2Name.Text = "CRIT DMG"
                cmbFlowerSubstat3Name.Text = "CRIT Rate"
                cmbFlowerSubstat4Name.Text = "ATK"
                txtFlowerHP.Text = "4780"
                txtFlowerSubstat1Value.Text = "9.3"
                txtFlowerSubstat2Value.Text = "10.9"
                txtFlowerSubstat3Value.Text = "6.6"
                txtFlowerSubstat4Value.Text = "51"

                cmbPlumeSet.Text = "Blizzard Strayer"
                cmbPlumeSubstat1Name.Text = "CRIT DMG"
                cmbPlumeSubstat2Name.Text = "CRIT Rate"
                cmbPlumeSubstat3Name.Text = "Energy Recharge"
                cmbPlumeSubstat4Name.Text = "DEF"
                txtPlumeATK.Text = "311"
                txtPlumeSubstat1Value.Text = "19.4"
                txtPlumeSubstat2Value.Text = "7.0"
                txtPlumeSubstat3Value.Text = "5.2"
                txtPlumeSubstat4Value.Text = "65"

                cmbSandsSet.Text = "Blizzard Strayer"
                cmbSandsMainStatName.Text = "ATK%"
                cmbSandsSubstat1Name.Text = "CRIT DMG"
                cmbSandsSubstat2Name.Text = "ATK"
                cmbSandsSubstat3Name.Text = "DEF"
                cmbSandsSubstat4Name.Text = "CRIT Rate"
                txtSandsMainStatValue.Text = "46.6"
                txtSandsSubstat1Value.Text = "20.2"
                txtSandsSubstat2Value.Text = "18"
                txtSandsSubstat3Value.Text = "19"
                txtSandsSubstat4Value.Text = "8.6"

                cmbGobletSet.Text = "Emblem of Severed Fates"
                cmbGobletMainStatName.Text = "Cryo DMG Bonus"
                cmbGobletSubstat1Name.Text = "CRIT DMG"
                cmbGobletSubstat2Name.Text = "CRIT Rate"
                cmbGobletSubstat3Name.Text = "ATK"
                cmbGobletSubstat4Name.Text = "ATK%"
                txtGobletMainStatValue.Text = "46.6"
                txtGobletSubstat1Value.Text = "20.2"
                txtGobletSubstat2Value.Text = "7.0"
                txtGobletSubstat3Value.Text = "14"
                txtGobletSubstat4Value.Text = "12.8"

                cmbCircletSet.Text = "Blizzard Strayer"
                cmbCircletMainStatName.Text = "CRIT DMG"
                cmbCircletSubstat1Name.Text = "DEF%"
                cmbCircletSubstat2Name.Text = "ATK%"
                cmbCircletSubstat3Name.Text = "HP%"
                cmbCircletSubstat4Name.Text = "ATK"
                txtCircletMainStatValue.Text = "62.2"
                txtCircletSubstat1Value.Text = "6.6"
                txtCircletSubstat2Value.Text = "14.0"
                txtCircletSubstat3Value.Text = "10.5"
                txtCircletSubstat4Value.Text = "31"
        End Select
        Call CalculateStats()
    End Sub
    'this displays the refinement bonus on the edit weapon screen
    Private Sub RefinementBonusChecker(ByVal strRefineBonus As String)
        Select Case strWeaponName
            Case "Mistsplitter Reforged"
                lblRefineBonus.Text = "Gain a " + strRefineBonus + "% Elemental DMG Bonus"
            Case "Staff of Homa"
                lblRefineBonus.Text = "HP increased by " + strRefineBonus + "%"
            Case "Song of Broken Pines"
                lblRefineBonus.Text = "Increases ATK by " + strRefineBonus + "%"
            Case "Wolf's Gravestone"
                lblRefineBonus.Text = "Increases ATK by " + strRefineBonus + "%"
            Case "Everlasting Moonglow"
                lblRefineBonus.Text = "Healing Bonus increased by " + strRefineBonus + "%"
            Case Else
                lblRefineBonus.Text = "This weapon has no passive refinement bonuses"
        End Select
    End Sub

    Private Sub btnClearArtifacts_Click(sender As Object, e As EventArgs) Handles btnClearArtifacts.Click
        cmbFlowerSet.Text = String.Empty
        cmbFlowerSubstat1Name.Text = String.Empty
        cmbFlowerSubstat2Name.Text = String.Empty
        cmbFlowerSubstat3Name.Text = String.Empty
        cmbFlowerSubstat4Name.Text = String.Empty
        txtFlowerHP.Text = String.Empty
        txtFlowerSubstat1Value.Text = String.Empty
        txtFlowerSubstat2Value.Text = String.Empty
        txtFlowerSubstat3Value.Text = String.Empty
        txtFlowerSubstat4Value.Text = String.Empty

        cmbPlumeSet.Text = String.Empty
        cmbPlumeSubstat1Name.Text = String.Empty
        cmbPlumeSubstat2Name.Text = String.Empty
        cmbPlumeSubstat3Name.Text = String.Empty
        cmbPlumeSubstat4Name.Text = String.Empty
        txtPlumeATK.Text = String.Empty
        txtPlumeSubstat1Value.Text = String.Empty
        txtPlumeSubstat2Value.Text = String.Empty
        txtPlumeSubstat3Value.Text = String.Empty
        txtPlumeSubstat4Value.Text = String.Empty

        cmbSandsSet.Text = String.Empty
        cmbSandsMainStatName.Text = String.Empty
        cmbSandsSubstat1Name.Text = String.Empty
        cmbSandsSubstat2Name.Text = String.Empty
        cmbSandsSubstat3Name.Text = String.Empty
        cmbSandsSubstat4Name.Text = String.Empty
        txtSandsMainStatValue.Text = String.Empty
        txtSandsSubstat1Value.Text = String.Empty
        txtSandsSubstat2Value.Text = String.Empty
        txtSandsSubstat3Value.Text = String.Empty
        txtSandsSubstat4Value.Text = String.Empty

        cmbGobletSet.Text = String.Empty
        cmbGobletMainStatName.Text = String.Empty
        cmbGobletSubstat1Name.Text = String.Empty
        cmbGobletSubstat2Name.Text = String.Empty
        cmbGobletSubstat3Name.Text = String.Empty
        cmbGobletSubstat4Name.Text = String.Empty
        txtGobletMainStatValue.Text = String.Empty
        txtGobletSubstat1Value.Text = String.Empty
        txtGobletSubstat2Value.Text = String.Empty
        txtGobletSubstat3Value.Text = String.Empty
        txtGobletSubstat4Value.Text = String.Empty

        cmbCircletSet.Text = String.Empty
        cmbCircletMainStatName.Text = String.Empty
        cmbCircletSubstat1Name.Text = String.Empty
        cmbCircletSubstat2Name.Text = String.Empty
        cmbCircletSubstat3Name.Text = String.Empty
        cmbCircletSubstat4Name.Text = String.Empty
        txtCircletMainStatValue.Text = String.Empty
        txtCircletSubstat1Value.Text = String.Empty
        txtCircletSubstat2Value.Text = String.Empty
        txtCircletSubstat3Value.Text = String.Empty
        txtCircletSubstat4Value.Text = String.Empty
    End Sub

    Private Sub txtFlowerSubstats_TextChanged(sender As Object, e As KeyPressEventArgs) Handles txtFlowerSubstat4Value.KeyPress, txtFlowerSubstat3Value.KeyPress, txtFlowerSubstat2Value.KeyPress, txtFlowerSubstat1Value.KeyPress, txtFlowerHP.KeyPress, txtPlumeSubstat4Value.KeyPress, txtPlumeSubstat3Value.KeyPress, txtPlumeSubstat2Value.KeyPress, txtPlumeSubstat1Value.KeyPress, txtSandsSubstat4Value.KeyPress, txtSandsSubstat3Value.KeyPress, txtSandsSubstat2Value.KeyPress, txtSandsSubstat1Value.KeyPress, txtSandsMainStatValue.KeyPress, txtPlumeATK.KeyPress, txtGobletSubstat4Value.KeyPress, txtGobletSubstat3Value.KeyPress, txtGobletSubstat2Value.KeyPress, txtGobletSubstat1Value.KeyPress, txtGobletMainStatValue.KeyPress, txtCircletSubstat4Value.KeyPress, txtCircletSubstat3Value.KeyPress, txtCircletSubstat2Value.KeyPress, txtCircletSubstat1Value.KeyPress, txtCircletMainStatValue.KeyPress
        If Not Char.IsNumber(e.KeyChar) And Not e.KeyChar = Chr(Keys.Delete) And Not e.KeyChar = Chr(Keys.Back) And Not e.KeyChar = Chr(Keys.Space) And e.KeyChar <> "." Then
            e.Handled = True
            Dim msg As String
            msg = "This field will accept only numbers and decimal system."
            MessageBox.Show(msg, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class

'------------------------------------------------------------
'          .-           .             .-. .-._.---'    .      
'  .---;`-'            /   .--.    .-/ -'(_) /  .-.   /       
' (   (_) .  .-.  .-../   /    )`-'-/--     /--.`-'  /   .-.  
'  )--     )/   )(   /   /    /    /       /   /    /  ./.-'_ 
' (      /'/   (  `-'-..(    /  `.'     .-/ _.(__._/_.-(__.'  
' `\___.'       `-       `-.'          (_/               
'--------------------------------------------------------
