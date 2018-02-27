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

#Hide-Powershell #Powershell ablak elréjtése
$sync = [Hashtable]::Synchronized(@{}) #for talking across runspaces.

 ######## függvények blokk kezdete ###########

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
    $Script:file=Save-File $PSScriptRoot # variable scope for the whole script
    $mSavePath.Text="Mentési útvonal: $($file)"
    $sync.file=$file
    if ($mConnect.Enabled -eq $false)
    {$mStart.Enabled = $true}    
}

Function Connect_Devices 
{
    if ($mTapegyseg_combobox.SelectedItem -eq $null -or $mPicoammeter_combobox.SelectedItem -eq $null)
    {
        $mConnect_Label.ForeColor = [System.Drawing.Color]::Red
        $mConnect_Label.Text="Add meg a COM portokat!"
    }
    else
    {
        $GWCOMport = $mTapegyseg_combobox.SelectedItem -as [string]
        $Script:gwport= new-Object System.IO.Ports.SerialPort $GWCOMport,9600,None,8,one # variable scope for the whole script
        $K6485COMport = $mPicoammeter_combobox.SelectedItem -as [string]
        $Script:k6485port= new-Object System.IO.Ports.SerialPort $K6485COMport,9600,None,8,one # variable scope for the whole script
        try
        {
            $gwport.Open()
            $sync.gwport=$gwport
            $k6485port.Open()
            $sync.k6485port=$k6485port
            $mConnect_Label.Text="A csatlakozas sikeres"
            $gwport.WriteLine("VSET2:12")
            $gwport.WriteLine("ISET2:0.5")
            $gwport.WriteLine("VSET1:3.8")
            $gwport.WriteLine("ISET1:0")
            $gwport.WriteLine("OUT1")
            $k6485port.WriteLine("*RST")
            $k6485port.WriteLine("SYST:ZCH ON")
            $k6485port.WriteLine("RANG:AUTO ON")
            $k6485port.WriteLine("SYST:ZCH OFF")
            $mConnect.Enabled = $false         
        }
        catch
        {
          $mConnect_Label.Text="A port foglalt vagy egyéb hiba történt"
        }
    }
 }

# long running task.
$PID_control = {
$PID_c = [PowerShell]::Create().AddScript({ # a tenyleges munkátvégző kód
    
    function Measure-detector_current #must be inside the runspace, custom functions outside of the runspace cannot be called?
    {
        $sync.k6485port.WriteLine("READ?")
        $PhotoCurrentString = ($sync.k6485port.ReadLine()).Trimend()
        [double]$PhotoCurrent=$PhotoCurrentString.Substring(0,$PhotoCurrentString.IndexOf("A")) #string-double konverzió
        [double]$PhotoCurrent=[double]$PhotoCurrent*-1
        return [double]$PhotoCurrent
    }
    
    $sync.mStart.Enabled = $false
    $sync.mStop.Enabled = $true
    
    #[double]$set_point = $sync.mCel_Label.Text/1000000000
    #[double]$KP = $sync.mKP_Label.Text
    #[double]$KI = $sync.mKI_Label.Text
    #[double]$KD = $sync.mKD_Label.Text
    #double]$iteration_time = $sync.mMintavet_Label.Text
    
    [double]$error_difference = 0
    [double]$actual_value = 0
    [double]$derivative = 0
    [double]$output = 0

    [double]$error_prior = 0
    [double]$integral = 0

    [double]$intergral_prior = 0

    [double]$m=3884874
    [double]$b=-0.1254
    #$i=0 #debug variable
 :labeled_loop While ($true) #végtelen ciklus
    {
        $actual_value = Measure-detector_current #detektoráram mérése

        $error_difference = $([double]$sync.mCel_Label.Text/1000000000) - $actual_value
        $integral = $integral + ($error_difference * [double]$sync.mMintavet_Label.Text)
            if ($integral -lt 0)
            {
                $integral = 0
            }
            if ($output -eq 1)
            {
                $integral = $intergral_prior
            }
            if ($output -eq 1 -and ($error_difference -gt $error_prior * 1.1 -or $error_difference -lt $error_prior * 0.9))
            {
                $integral = $integral + ($error_difference * [double]$sync.mMintavet_Label.Text)
            }
        $derivative = ($error_difference - $error_prior)/[double]$sync.mMintavet_Label.Text
        $output = [double]$sync.mKP_Label.Text * $error_difference + [double]$sync.mKI_Label.Text * $integral + [double]$sync.mKD_Label.Text * $derivative
        $output = $m * $output + $b #detektor áram konvertálása LED meghajtóárammá
        $output =  [System.Math]::Round($output,3) #LED meghajtóáram konvetálása 3 tizedesjegyre
        if ($output -lt 0)
            {
                $output = 0
            }
            elseif ($output -gt 1)
            {
                $output = 1
            }
        $sync.gwport.WriteLine("ISET1:$($output)") #tápegységnek az új LED meghajtóáram érték küldése

        $([double]$sync.mCel_Label.Text/1000000000).ToString() + "`t" + $actual_value.ToString() + "`t" + $error_difference.ToString() + "`t" + $integral.ToString() + "`t" + $derivative.ToString() + "`t" + $output.ToString() | Out-File $sync.file -Append

        $error_prior  = $error_difference
        $intergral_prior = $integral
        
        Start-Sleep -Milliseconds $([double]$sync.mMintavet_Label.Text)

        if ($sync.mStop.Enabled -eq $false)
        {
            break labeled_loop 
        }

        <#debug
        $sync.mConnect_Label.Text = "OUT$($i % 2)"
        $sync.gwport.WriteLine("OUT$($i % 2)")
        if ($sync.mStop.Enabled -eq $false)
            {
            break labeled_loop 
            }
        $i++
        Start-Sleep -Milliseconds 2000#>

     }
     $sync.k6485port.Close()
     $sync.gwport.Close()
     $sync.mStart.Enabled = $true
     $sync.mConnect.Enabled = $true
     $sync.mConnect_Label.Text = "Műszerek lecsatlakoztatva"    
})

$runspace = [RunspaceFactory]::CreateRunspace() #Creates a single runspace that uses the default host and runspace configuration.
$runspace.ApartmentState = "STA" #Single Threaded Apartment (STA) thread-safe
$runspace.ThreadOptions = "ReuseThread" #Creates a new thread for the first invocation and then re-uses that thread in subsequent invocations.
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("sync", $sync) #Sharing Variables and Live Objects Between PowerShell Runspaces

$PID_c.Runspace = $runspace # Add $PID_c to runspace
$PID_c.BeginInvoke()
}

$StopPIDloop = {
    $sync.mStop.Enabled = $false
    $sync.period = 1
}

######## függvények blokk vége ###########

############ GUI blokk kezdete #########################     
    Add-Type -AssemblyName System.Windows.Forms 
    Add-Type -AssemblyName System.Drawing 
    $MyForm = New-Object System.Windows.Forms.Form 
    $MyForm.Text="Fény szabályzás" 
    $MyForm.Size = New-Object System.Drawing.Size(500,750) 
     
#tapegység kiválasztása combo 
        $mTapegyseg_combobox = New-Object System.Windows.Forms.ComboBoX 
                $mTapegyseg_combobox.Text="" 
                $mTapegyseg_combobox.Top="50" 
                $mTapegyseg_combobox.Left="49" 
                $mTapegyseg_combobox.Anchor="Left,Top" 
        $mTapegyseg_combobox.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mTapegyseg_combobox) 
        $mTapegyseg_combobox.Items.AddRange([System.IO.Ports.SerialPort]::getportnames()) 
#tapegység kiválasztása label
        $mTapegyseg_Label = New-Object System.Windows.Forms.Label 
                $mTapegyseg_Label.Text="Tápegység port" 
                $mTapegyseg_Label.Top="25" 
                $mTapegyseg_Label.Left="54" 
                $mTapegyseg_Label.Anchor="Left,Top" 
        $mTapegyseg_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mTapegyseg_Label) 
         
#picoammeter kiválasztása combo  
        $mPicoammeter_combobox = New-Object System.Windows.Forms.ComboBoX 
                $mPicoammeter_combobox.Text="" 
                $mPicoammeter_combobox.Top="51" 
                $mPicoammeter_combobox.Left="188" 
                $mPicoammeter_combobox.Anchor="Left,Top" 
        $mPicoammeter_combobox.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mPicoammeter_combobox) 
        $mPicoammeter_combobox.Items.AddRange([System.IO.Ports.SerialPort]::getportnames()) 
#picoammeter kiválasztása label
        $mPicoammeter_Label = New-Object System.Windows.Forms.Label 
                $mPicoammeter_Label.Text="Picoampermérő port" 
                $mPicoammeter_Label.Top="25" 
                $mPicoammeter_Label.Left="185" 
                $mPicoammeter_Label.Anchor="Left,Top" 
        $mPicoammeter_Label.Size = New-Object System.Drawing.Size(120,23) 
        $MyForm.Controls.Add($mPicoammeter_Label)

#csatlakozás button label
        $mConnect_Label = New-Object System.Windows.Forms.Label 
                $mConnect_Label.Text="Válaszd ki a műszerek COM portját és nyomd meg a csatlakozás gombot." 
                $mConnect_Label.Top="15" 
                $mConnect_Label.Left="320" 
                $mConnect_Label.Anchor="Left,Top" 
        $mConnect_Label.Size = New-Object System.Drawing.Size(150,50) 
        $MyForm.Controls.Add($mConnect_Label) 
         
#csatalakozás button
        $mConnect = New-Object System.Windows.Forms.Button 
                $mConnect.Text="Csatlakozás" 
                $mConnect.Top="70" 
                $mConnect.Left="350" 
                $mConnect.Anchor="Left,Top" 
        $mConnect.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mConnect) 
         
#cel track           
        $mCel_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mCel_TrackBar.Text="TrackBar1" 
                $mCel_TrackBar.Top="108" 
                $mCel_TrackBar.Left="22" 
                $mCel_TrackBar.Anchor="Left,Top"

                $mCel_TrackBar.SetRange(10,160)
                $mCel_TrackBar.TickFrequency=5
                $mCel_TrackBar.Value=140
                $TrackLabel_CEL_Value=140

        $mCel_TrackBar.Size = New-Object System.Drawing.Size(460,23)

        $mCel_TrackBar.add_ValueChanged({
        $TrackLabel_CEL_Value = $mCel_TrackBar.Value
        $mCel_Label.Text = $TrackLabel_CEL_Value #"Célérték: $($TrackLabel_CEL_Value)nA"
        })

        $MyForm.Controls.Add($mCel_TrackBar) 
         
#cel label
        $mCel_Label = New-Object System.Windows.Forms.Label 
                $mCel_Label.Text= $TrackLabel_CEL_Value #"Célérték: $($TrackLabel_CEL_Value)nA"
                $mCel_Label.Top="162" 
                $mCel_Label.Left="218" 
                $mCel_Label.Anchor="Left,Top" 
        $mCel_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mCel_Label) 
         
#P track
        $mKP_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKP_TrackBar.Text="TrackBar2" 
                $mKP_TrackBar.Top="195" 
                $mKP_TrackBar.Left="20" 
                $mKP_TrackBar.Anchor="Left,Top"

                 $mKP_TrackBar.SetRange(10,1100)
                 $mKP_TrackBar.TickFrequency=50
                 $mKP_TrackBar.Value=100
                $TrackLabel_P_Value=100

        $mKP_TrackBar.Size = New-Object System.Drawing.Size(460,23)

          $mKP_TrackBar.add_ValueChanged({
        $TrackLabel_P_Value =  $mKP_TrackBar.Value
        $mKP_Label.Text = $TrackLabel_P_Value/1000 #"KP értéke: $($TrackLabel_P_Value/1000)"
        })

        $MyForm.Controls.Add($mKP_TrackBar) 
         
#P label
        $mKP_Label = New-Object System.Windows.Forms.Label 
                $mKP_Label.Text= $TrackLabel_P_Value/1000 #"KP értéke: $($TrackLabel_P_Value/1000)"
                $mKP_Label.Top="252" 
                $mKP_Label.Left="217" 
                $mKP_Label.Anchor="Left,Top" 
        $mKP_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKP_Label) 
         
#I track
        $mKI_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKI_TrackBar.Text="TrackBar3" 
                $mKI_TrackBar.Top="289" 
                $mKI_TrackBar.Left="21" 
                $mKI_TrackBar.Anchor="Left,Top"

                $mKI_TrackBar.SetRange(1,110)
                $mKI_TrackBar.TickFrequency=5
                $mKI_TrackBar.Value=10
                $TrackLabel_I_Value=10

        $mKI_TrackBar.Size = New-Object System.Drawing.Size(460,23) 
        
        $mKI_TrackBar.add_ValueChanged({
        $TrackLabel_I_Value =  $mKI_TrackBar.Value
        $mKI_Label.Text = $TrackLabel_I_Value/1000 #"KI értéke: $($TrackLabel_I_Value/1000)"
        })

        $MyForm.Controls.Add($mKI_TrackBar) 
         
#I label
        $mKI_Label = New-Object System.Windows.Forms.Label 
                $mKI_Label.Text= $TrackLabel_I_Value/1000 #"KI értéke: $($TrackLabel_I_Value/1000)"
                $mKI_Label.Top="346" 
                $mKI_Label.Left="223" 
                $mKI_Label.Anchor="Left,Top" 
        $mKI_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKI_Label) 
         
#D trackbar
        $mKD_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mKD_TrackBar.Text="TrackBar4" 
                $mKD_TrackBar.Top="392" 
                $mKD_TrackBar.Left="23" 
                $mKD_TrackBar.Anchor="Left,Top"

                $mKD_TrackBar.SetRange(0,1000)
                $mKD_TrackBar.TickFrequency=50
                $mKD_TrackBar.Value=0
                $TrackLabel_D_Value=0

        $mKD_TrackBar.Size = New-Object System.Drawing.Size(460,23) 

        $mKD_TrackBar.add_ValueChanged({
        $TrackLabel_D_Value =  $mKD_TrackBar.Value
        $mKD_Label.Text = $TrackLabel_D_Value/1000 #"KD értéke: $($TrackLabel_D_Value/1000)"
        })

        $MyForm.Controls.Add($mKD_TrackBar) 
         
#D label
        $mKD_Label = New-Object System.Windows.Forms.Label 
                $mKD_Label.Text= $TrackLabel_D_Value/1000 #"KD értéke: $($TrackLabel_D_Value/1000)"
                $mKD_Label.Top="449" 
                $mKD_Label.Left="221" 
                $mKD_Label.Anchor="Left,Top" 
        $mKD_Label.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mKD_Label) 
         
#sample rate track 
        $mMintavet_TrackBar = New-Object System.Windows.Forms.TrackBar 
                $mMintavet_TrackBar.Text="TrackBar5" 
                $mMintavet_TrackBar.Top="492" 
                $mMintavet_TrackBar.Left="24" 
                $mMintavet_TrackBar.Anchor="Left,Top"
                
                $mMintavet_TrackBar.SetRange(70,7000)
                $mMintavet_TrackBar.TickFrequency=70
                $mMintavet_TrackBar.Value=70
                $TrackLabel_M_Value=70
                 
        $mMintavet_TrackBar.Size = New-Object System.Drawing.Size(460,23)

        $mMintavet_TrackBar.add_ValueChanged({
        $TrackLabel_M_Value =  $mMintavet_TrackBar.Value
        $mMintavet_Label.Text = $TrackLabel_M_Value #"Mintavételezési idő: $($TrackLabel_M_Value)ms"
        })

        $MyForm.Controls.Add($mMintavet_TrackBar) 
         
#sample rate label
        $mMintavet_Label = New-Object System.Windows.Forms.Label 
                $mMintavet_Label.Text= $TrackLabel_M_Value #"Mintavételezési idő: $($TrackLabel_M_Value)ms"
                $mMintavet_Label.Top="547" 
                $mMintavet_Label.Left="200" 
                $mMintavet_Label.Anchor="Left,Top" 
        $mMintavet_Label.Size = New-Object System.Drawing.Size(120,30) 
        $MyForm.Controls.Add($mMintavet_Label) 
         
#mentési útvonal label 
        $mSavePath = New-Object System.Windows.Forms.Label 
                $mSavePath.Text="Mentési útvonal: " 
                $mSavePath.Top="587" 
                $mSavePath.Left="28" 
                $mSavePath.Anchor="Left,Top" 
        $mSavePath.Size = New-Object System.Drawing.Size(460,50) 
        $MyForm.Controls.Add($mSavePath) 
         
 #mentési útvonal button
        $mSave = New-Object System.Windows.Forms.Button 
                $mSave.Text="Mentési útvonal" 
                $mSave.Top="646" 
                $mSave.Left="37" 
                $mSave.Anchor="Left,Top" 
        $mSave.Size = New-Object System.Drawing.Size(100,23) 
        $MyForm.Controls.Add($mSave) 
         
#start PID control button
        $mStart = New-Object System.Windows.Forms.Button 
                $mStart.Text="Indítás" 
                $mStart.Top="647" 
                $mStart.Left="207" 
                $mStart.Anchor="Left,Top" 
        $mStart.Size = New-Object System.Drawing.Size(100,23)
        $mStart.Enabled = $false
        $MyForm.Controls.Add($mStart) 
         
#Stop PID control button
        $mStop = New-Object System.Windows.Forms.Button 
                $mStop.Text="Leállítás" 
                $mStop.Top="647" 
                $mStop.Left="366" 
                $mStop.Anchor="Left,Top" 
        $mStop.Size = New-Object System.Drawing.Size(100,23)
        $mStop.Enabled = $false
        $MyForm.Controls.Add($mStop) 
        
        $mSave.Add_Click({Export_to_txt})
        $mConnect.Add_Click({Connect_Devices})
        $mStart.Add_Click($PID_control)
        $mStop.Add_Click($StopPIDloop)

        # add controls to the form.
        $sync.mConnect = $mConnect
        $sync.mConnect_Label = $mConnect_Label
        $sync.mStart = $mStart
        $sync.mStop = $mStop
        $sync.mCel_Label = $mCel_Label     
        $sync.mKP_Label = $mKP_Label 
        $sync.mKI_Label = $mKI_Label
        $sync.mKD_Label = $mKD_Label
        $sync.mMintavet_Label = $mMintavet_Label

        $MyForm.Controls.AddRange(@($sync.mConnect, $sync.mConnect_Label, $sync.mStart, $sync.mStop, $sync.mCel_Label, $sync.mKP_Label, $sync.mKI_Label, $sync.mKD_Label, $sync.mMintavet_Label))

        $MyForm.ShowDialog()  

############ GUI blokk vége ######################### 
