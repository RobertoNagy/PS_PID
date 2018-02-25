######### Függvények importálása C++ DLL-ből blokk kezdete ############ 
$Script:showWindowAsync = Add-Type -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru
Function Show-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 10)
}
Function Hide-Powershell()
{
$null = $showWindowAsync::ShowWindowAsync((Get-Process -Id $pid).MainWindowHandle, 2)
}
######### Függvények importálása C++ DLL-ből blokk vége ############ 

Hide-Powershell #Powershell ablak elréjtése
 
 ######## függvények blokk kezdete ###########
function Measure-detector_current
{
    $portK6485.WriteLine("READ?")
    $PhotoCurrentString = ($portK6485.ReadLine()).Trimend()
    [double]$PhotoCurrent=$PhotoCurrentString.Substring(0,$PhotoCurrentString.IndexOf("A")) #string-double konverzió
    [double]$PhotoCurrent=[double]$PhotoCurrent*-1
    return [double]$PhotoCurrent
}

function Save-File([string] $initialDirectory ) 
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
 
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "TXT fajlok (*.txt)| *.txt"
    $OpenFileDialog.OverwritePrompt = $false
    $OpenFileDialog.ShowDialog() |  Out-Null
 
    return $OpenFileDialog.filename
}

function Export_to_txt
{
    $file=Save-File $PSScriptRoot
    $Script:UIArray | out-file $file #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
}

Function Connect_Devices 
{
if ($mComboBoX1.SelectedItem -eq $null -or $mComboBoX2.SelectedItem -eq $null)
{
$mLabel11.Text="Add meg a COM portokat!"
}
else
{
$GWCOMport = $mComboBoX1.SelectedItem -as [string]
$Script:gwport= new-Object System.IO.Ports.SerialPort $GWCOMport,9600,None,8,one
$K2000COMport = $mComboBoX2.SelectedItem -as [string]
$Script:k2000port= new-Object System.IO.Ports.SerialPort $K2000COMport,9600,None,8,one
try
{
$Script:gwport.Open()
$Script:k2000port.Open()
$mLabel11.Text="A csatlakozas sikeres"
$Script:k2000port.WriteLine("SYSTem:BEEPer:STATe 0")
$Script:k2000port.WriteLine(":INITiate:CONTinuous OFF")
$mButton2.Enabled = $true
}
catch
{
$mLabel11.Text="A port foglalt vagy egyéb hiba történt"
}
}
 
}

######## függvények blokk vége ###########

############ GUI blokk kezdete #########################     
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="Fény szabályzás" 
    $MyForm.Size = New-Object System.Drawing.Size(500,750) 
     
 
        $mTapegyseg_combobox = New-Object System.Windows.Forms.ComboBoX 
                $mTapegyseg_combobox.Text="" 
                $mTapegyseg_combobox.Top="50" 
                $mTapegyseg_combobox.Left="49" 
                $mTapegyseg_combobox.Anchor="Left,Top" 
        $mTapegyseg_combobox.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mTapegyseg_combobox) 
        $mTapegyseg_combobox.Items.AddRange([System.IO.Ports.SerialPort]::getportnames()) 
 
        $mTapegyseg_Label = New-Object System.Windows.Forms.Label 
                $mTapegyseg_Label.Text="Tápegység port" 
                $mTapegyseg_Label.Top="25" 
                $mTapegyseg_Label.Left="54" 
                $mTapegyseg_Label.Anchor="Left,Top" 
        $mTapegyseg_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mTapegyseg_Label) 
         
 
        $mPicoammeter_combobox = New-Object System.Windows.Forms.ComboBoX 
                $mPicoammeter_combobox.Text="" 
                $mPicoammeter_combobox.Top="51" 
                $mPicoammeter_combobox.Left="188" 
                $mPicoammeter_combobox.Anchor="Left,Top" 
        $mPicoammeter_combobox.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mPicoammeter_combobox) 
        $mPicoammeter_combobox.Items.AddRange([System.IO.Ports.SerialPort]::getportnames()) 
 
        $mPicoammeter_Label = New-Object System.Windows.Forms.Label 
                $mPicoammeter_Label.Text="Picoampermérő port" 
                $mPicoammeter_Label.Top="25" 
                $mPicoammeter_Label.Left="185" 
                $mPicoammeter_Label.Anchor="Left,Top" 
        $mPicoammeter_Label.Size = New-Object System.Drawing.Size(120,23) 
        $MyForm.Controls.Add($mPicoammeter_Label) 
         
 
        $mConnect = New-Object System.Windows.Forms.Button 
                $mConnect.Text="Csatlakozás" 
                $mConnect.Top="51" 
                $mConnect.Left="349" 
                $mConnect.Anchor="Left,Top" 
        $mConnect.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mConnect) 
         
 
        $mCel_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mCel_TrackBar.Text="TrackBar1" 
                $mCel_TrackBar.Top="108" 
                $mCel_TrackBar.Left="22" 
                $mCel_TrackBar.Anchor="Left,Top" 
        $mCel_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        $MyForm.Controls.Add($mCel_TrackBar) 
         
 
        $mCel_Label = New-Object System.Windows.Forms.Label 
                $mCel_Label.Text="Célérték" 
                $mCel_Label.Top="162" 
                $mCel_Label.Left="218" 
                $mCel_Label.Anchor="Left,Top" 
        $mCel_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mCel_Label) 
         
 
        $mKP_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKP_TrackBar.Text="TrackBar2" 
                $mKP_TrackBar.Top="195" 
                $mKP_TrackBar.Left="20" 
                $mKP_TrackBar.Anchor="Left,Top" 
        $mKP_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        $MyForm.Controls.Add($mKP_TrackBar) 
         
 
        $mKP_Label = New-Object System.Windows.Forms.Label 
                $mKP_Label.Text="KP értéke" 
                $mKP_Label.Top="252" 
                $mKP_Label.Left="217" 
                $mKP_Label.Anchor="Left,Top" 
        $mKP_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKP_Label) 
         
 
        $mKI_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKI_TrackBar.Text="TrackBar3" 
                $mKI_TrackBar.Top="289" 
                $mKI_TrackBar.Left="21" 
                $mKI_TrackBar.Anchor="Left,Top" 
        $mKI_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        $MyForm.Controls.Add($mKI_TrackBar) 
         
 
        $mKI_Label = New-Object System.Windows.Forms.Label 
                $mKI_Label.Text="KI értéke" 
                $mKI_Label.Top="346" 
                $mKI_Label.Left="223" 
                $mKI_Label.Anchor="Left,Top" 
        $mKI_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKI_Label) 
         
 
        $mKD_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKD_TrackBar.Text="TrackBar4" 
                $mKD_TrackBar.Top="392" 
                $mKD_TrackBar.Left="23" 
                $mKD_TrackBar.Anchor="Left,Top" 
        $mKD_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        $MyForm.Controls.Add($mKD_TrackBar) 
         
 
        $mKD_Label = New-Object System.Windows.Forms.Label 
                $mKD_Label.Text="KD értéke" 
                $mKD_Label.Top="449" 
                $mKD_Label.Left="221" 
                $mKD_Label.Anchor="Left,Top" 
        $mKD_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKD_Label) 
         
 
        $mMintavet_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mMintavet_TrackBar.Text="TrackBar5" 
                $mMintavet_TrackBar.Top="492" 
                $mMintavet_TrackBar.Left="24" 
                $mMintavet_TrackBar.Anchor="Left,Top" 
        $mMintavet_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        $MyForm.Controls.Add($mMintavet_TrackBar) 
         
 
        $mMintavet_Label = New-Object System.Windows.Forms.Label 
                $mMintavet_Label.Text="Mintavételezési idő" 
                $mMintavet_Label.Top="547" 
                $mMintavet_Label.Left="200" 
                $mMintavet_Label.Anchor="Left,Top" 
        $mMintavet_Label.Size = New-Object System.Drawing.Size(120,23) 
        $MyForm.Controls.Add($mMintavet_Label) 
         
 
        $mSavePath = New-Object System.Windows.Forms.Label 
                $mSavePath.Text="Mentési útvonal:" 
                $mSavePath.Top="587" 
                $mSavePath.Left="28" 
                $mSavePath.Anchor="Left,Top" 
        $mSavePath.Size = New-Object System.Drawing.Size(460,50) 
        $MyForm.Controls.Add($mSavePath) 
         
 
        $mSave = New-Object System.Windows.Forms.Button 
                $mSave.Text="Mentési útvonal" 
                $mSave.Top="646" 
                $mSave.Left="37" 
                $mSave.Anchor="Left,Top" 
        $mSave.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mSave) 
         
 
        $mStart = New-Object System.Windows.Forms.Button 
                $mStart.Text="Indítás" 
                $mStart.Top="647" 
                $mStart.Left="207" 
                $mStart.Anchor="Left,Top" 
        $mStart.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mStart) 
         
 
        $mStop = New-Object System.Windows.Forms.Button 
                $mStop.Text="Leállítás" 
                $mStop.Top="647" 
                $mStop.Left="366" 
                $mStop.Anchor="Left,Top" 
        $mStop.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mStop) 
        $MyForm.ShowDialog()

############ GUI blokk vége ######################### 
