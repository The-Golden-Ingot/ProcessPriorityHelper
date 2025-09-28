#Requires -Version 5.1
[CmdletBinding()]
param(
  [string]$LogPath,
  [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:LogPath = $LogPath
$script:Quiet   = $Quiet.IsPresent

# --------------------------
# Logging
# --------------------------
function Write-Log {
  param(
    [Parameter(Mandatory)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
  )
  $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $line = "[$timestamp][$Level] $Message"

  # Console
  if (-not $script:Quiet) {
    $color = switch ($Level) {
      'INFO' { 'Gray' }
      'WARN' { 'Yellow' }
      'ERROR' { 'Red' }
    }
    Write-Host $line -ForegroundColor $color
  }

  # File
  if ($script:LogPath) {
    try {
      $dir = Split-Path -LiteralPath $script:LogPath -Parent
      if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
      }
      Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    } catch {
      if (-not $script:Quiet) {
        Write-Host "[LOG][WARN] Failed to write to log file '$($script:LogPath)': $_" -ForegroundColor Yellow
      }
    }
  }
}

# --------------------------
# Elevation + 64-bit + STA + Desktop host relaunch
# --------------------------
function Test-IsAdmin {
  $id = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]$id
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Ensure-AdminAndUiHost {
  $selfPath   = $PSCommandPath
  $isPackaged = $false

  if (-not $selfPath) {
    try {
      $candidate = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
      if ($candidate -and (Test-Path -LiteralPath $candidate)) {
        $selfPath   = $candidate
        $isPackaged = $true
      }
    } catch {
      $selfPath = $null
    }
  }

  if (-not $selfPath) {
    throw "This script must be run from a file path (PSCommandPath not set)."
  }

  $isAdmin   = Test-IsAdmin
  $needAdmin = -not $isAdmin

  $is64OS    = [Environment]::Is64BitOperatingSystem
  $is64Proc  = [Environment]::Is64BitProcess
  $need64    = $is64OS -and -not $is64Proc

  $needSta   = ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA')

  $currentHostPath = (Get-Process -Id $PID).Path
  $hostLeaf        = [System.IO.Path]::GetFileName($currentHostPath).ToLowerInvariant()

  # Force a Desktop PowerShell host (powershell.exe) so we can guarantee -STA
  $preferDesktopPs = $false
  if (-not $isPackaged) {
    $preferDesktopPs = ($hostLeaf -ne 'powershell.exe')
  }

  $needRelaunch = $needAdmin -or $need64 -or $needSta -or $preferDesktopPs
  if (-not $needRelaunch) { return }

  # Build target host path (Desktop PowerShell)
  if ($isPackaged) {
    $targetHost = $selfPath
    $args = @()
  } elseif ($is64OS) {
    $targetHost = "$env:WINDIR\sysnative\WindowsPowerShell\v1.0\powershell.exe"
    $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA',"-File","`"$selfPath`"")
  } else {
    $targetHost = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
    $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA',"-File","`"$selfPath`"")
  }

  # Pass through supported parameters
  foreach ($k in $PSBoundParameters.Keys) {
    switch ($k) {
      'LogPath' { if ($PSBoundParameters[$k]) { $args += '-LogPath'; $args += "`"$($PSBoundParameters[$k])`"" } }
      'Quiet'   { if ($PSBoundParameters[$k]) { $args += '-Quiet' } }
    }
  }

  $args = @($args | Where-Object { $_ -ne $null })

  $why = @()
  if ($needAdmin)     { $why += 'elevated' }
  if ($need64)        { $why += '64-bit' }
  if ($needSta)       { $why += 'STA' }
  if ($preferDesktopPs){ $why += 'PowerShell.exe (Desktop)' }

  Write-Host "Relaunching ($($why -join ', ')): $targetHost" -ForegroundColor Yellow
  $sp = @{ FilePath = $targetHost }
  if ($args.Count -gt 0) { $sp['ArgumentList'] = $args }
  if ($needAdmin) { $sp['Verb'] = 'RunAs' }
  Start-Process @sp | Out-Null
  exit
}

Ensure-AdminAndUiHost

# --------------------------
# Assemblies (WPF)
# --------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --------------------------
# Constants and options
# --------------------------
$IFEO_Relative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'

$CpuOptions = @(
  [pscustomobject]@{ Name = 'Realtime (Use with extreme caution)'; Value = 4 },
  [pscustomobject]@{ Name = 'High'; Value = 3 },
  [pscustomobject]@{ Name = 'Above Normal'; Value = 6 },
  [pscustomobject]@{ Name = 'Normal'; Value = 2 },
  [pscustomobject]@{ Name = 'Below Normal'; Value = 5 },
  [pscustomobject]@{ Name = 'Idle'; Value = 1 }
)

$IoOptions = @(
  [pscustomobject]@{ Name = 'N/A'; Value = $null },
  [pscustomobject]@{ Name = 'Very Low'; Value = 0 },
  [pscustomobject]@{ Name = 'Low'; Value = 1 },
  [pscustomobject]@{ Name = 'Normal'; Value = 2 },
  [pscustomobject]@{ Name = 'High'; Value = 3 }
)

$PageOptions = @(
  [pscustomobject]@{ Name = 'N/A'; Value = $null },
  [pscustomobject]@{ Name = 'Low'; Value = 1 },
  [pscustomobject]@{ Name = 'Below Normal'; Value = 2 },
  [pscustomobject]@{ Name = 'Normal'; Value = 3 },
  [pscustomobject]@{ Name = 'Above Normal'; Value = 4 },
  [pscustomobject]@{ Name = 'High'; Value = 5 }
)

# --------------------------
# Utility
# --------------------------
function CpuValueToName([int]$value) {
  switch ($value) {
    1 { 'Idle' }
    2 { 'Normal' }
    3 { 'High' }
    4 { 'Realtime' }
    5 { 'Below Normal' }
    6 { 'Above Normal' }
    default { $value.ToString() }
  }
}

function IoValueToName([int]$value) {
  switch ($value) {
    0 { 'Very Low' }
    1 { 'Low' }
    2 { 'Normal' }
    3 { 'High' }
    default { $value.ToString() }
  }
}

function PageValueToName([int]$value) {
  switch ($value) {
    1 { 'Low' }
    2 { 'Below Normal' }
    3 { 'Normal' }
    4 { 'Above Normal' }
    5 { 'High' }
    default { $value.ToString() }
  }
}

function Show-ErrorMessage { param([string]$Message) [void][System.Windows.MessageBox]::Show($Message, 'Error', 'OK', 'Error') }
function Show-InfoMessage  { param([string]$Message) [void][System.Windows.MessageBox]::Show($Message, 'Information', 'OK', 'Information') }

# --------------------------
# Registry helpers (.NET API; 64-bit view on 64-bit OS)
# --------------------------
function Open-IFEOKey {
  param([bool]$Writable = $false)
  $view = if ([Environment]::Is64BitOperatingSystem) {
    [Microsoft.Win32.RegistryView]::Registry64
  } else {
    [Microsoft.Win32.RegistryView]::Default
  }
  $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
  $ifeo = $base.OpenSubKey($IFEO_Relative, $Writable)
  if (-not $ifeo -and $Writable) {
    $ifeo = $base.CreateSubKey($IFEO_Relative)
  }
  return $ifeo
}

function Ensure-PerfOptionsKey {
  param([Parameter(Mandatory)][string]$Executable)
  $ifeo = Open-IFEOKey -Writable $true
  if (-not $ifeo) { throw "Failed to open IFEO base key." }
  try {
    $exeKey = $ifeo.OpenSubKey($Executable, $true)
    if (-not $exeKey) { $exeKey = $ifeo.CreateSubKey($Executable) }
    $perfKey = $exeKey.OpenSubKey('PerfOptions', $true)
    if (-not $perfKey) { $perfKey = $exeKey.CreateSubKey('PerfOptions') }
    $exeKey.Close()
    return $perfKey
  } finally {
    $ifeo.Close()
  }
}

function Get-Overrides {
  $results = [System.Collections.ObjectModel.ObservableCollection[psobject]]::new()
  $ifeo = Open-IFEOKey -Writable $false
  if (-not $ifeo) { return $results }
  try {
    foreach ($exe in $ifeo.GetSubKeyNames()) {
      $perf = $ifeo.OpenSubKey("$exe\PerfOptions", $false)
      if (-not $perf) { continue }
      try {
        $cpu = $perf.GetValue('CpuPriorityClass', $null)
        $io  = $perf.GetValue('IoPriority', $null)
        $pg  = $perf.GetValue('PagePriority', $null)
        if ($cpu -eq $null -and $io -eq $null -and $pg -eq $null) { continue }
        $cpuVal = if ($cpu -ne $null) { [int]$cpu } else { $null }
        $ioVal  = if ($io  -ne $null) { [int]$io }  else { $null }
        $pgVal  = if ($pg  -ne $null) { [int]$pg }  else { $null }
        $results.Add([pscustomobject]@{
          Executable  = $exe
          CpuValue    = $cpuVal
          CpuDisplay  = if ($cpuVal -ne $null) { (CpuValueToName $cpuVal) } else { '' }
          IoValue     = $ioVal
          IoDisplay   = if ($ioVal  -ne $null) { (IoValueToName $ioVal) } else { '' }
          PageValue   = $pgVal
          PageDisplay = if ($pgVal -ne $null) { (PageValueToName $pgVal) } else { '' }
        })
      } finally {
        $perf.Close()
      }
    }
  } finally {
    $ifeo.Close()
  }

  if ($results.Count -le 1) { return $results }
  $sorted = [System.Collections.ObjectModel.ObservableCollection[psobject]]::new()
  foreach ($e in ($results | Sort-Object Executable)) { $sorted.Add($e) }
  return $sorted
}

function Apply-Override {
  param([Parameter(Mandatory)][pscustomobject]$Data)

  $exeRaw = $Data.Executable
  if ([string]::IsNullOrWhiteSpace($exeRaw)) { throw "Executable name is required." }
  $exeName = [System.IO.Path]::GetFileName($exeRaw.Trim())
  if (-not $exeName.EndsWith('.exe',[System.StringComparison]::OrdinalIgnoreCase)) { $exeName = "$exeName.exe" }

  $cpuValue  = [int]$Data.CpuValue
  $ioValue   = if ($Data.IoValue   -ne $null -and $Data.IoValue   -ne '') { [int]$Data.IoValue }   else { $null }
  $pageValue = if ($Data.PageValue -ne $null -and $Data.PageValue -ne '') { [int]$Data.PageValue } else { $null }

  if ($cpuValue -lt 1 -or $cpuValue -gt 6) { throw "CPU priority value '$cpuValue' is out of range." }

  Write-Log "Applying override for '$exeName' (CPU=$cpuValue, IO=$ioValue, Page=$pageValue)" 'INFO'

  $perfKey = Ensure-PerfOptionsKey -Executable $exeName
  try {
    $perfKey.SetValue('CpuPriorityClass', $cpuValue, [Microsoft.Win32.RegistryValueKind]::DWord)
    if ($ioValue -ne $null) { $perfKey.SetValue('IoPriority', $ioValue, [Microsoft.Win32.RegistryValueKind]::DWord) }
    else { try { $perfKey.DeleteValue('IoPriority', $false) } catch {} }

    if ($pageValue -ne $null) { $perfKey.SetValue('PagePriority', $pageValue, [Microsoft.Win32.RegistryValueKind]::DWord) }
    else { try { $perfKey.DeleteValue('PagePriority', $false) } catch {} }
  } finally {
    $perfKey.Close()
  }
}

function Remove-OverrideValues {
  param([Parameter(Mandatory)][string]$Executable)
  $exeName = [System.IO.Path]::GetFileName($Executable.Trim())
  $ifeo = Open-IFEOKey -Writable $true
  if (-not $ifeo) { return $false }
  try {
    $perf = $ifeo.OpenSubKey("$exeName\PerfOptions", $true)
    if (-not $perf) { return $false }
    try {
      foreach ($name in 'CpuPriorityClass','IoPriority','PagePriority') {
        try { $perf.DeleteValue($name, $false) } catch {}
      }
      return $true
    } finally {
      $perf.Close()
    }
  } finally {
    $ifeo.Close()
  }
}

function Remove-PerfOptionsKey {
  param([Parameter(Mandatory)][string]$Executable)
  $exeName = [System.IO.Path]::GetFileName($Executable.Trim())
  $ifeo = Open-IFEOKey -Writable $true
  if (-not $ifeo) { return $false }
  try {
    try {
      $ifeo.DeleteSubKeyTree($exeName, $false)
      return $true
    } catch {
      try {
        $exeKey = $ifeo.OpenSubKey($exeName, $true)
        if (-not $exeKey) { return $false }
        try {
          $exeKey.DeleteSubKeyTree('PerfOptions', $false)
        } finally {
          $exeKey.Close()
        }

        try {
          $ifeo.DeleteSubKey($exeName, $false)
        } catch {
          # If other values remain, leave the key intact.
        }
        return $true
      } catch {
        return $false
      }
    }
  } finally {
    $ifeo.Close()
  }
}

# --------------------------
# WPF helpers
# --------------------------
function Load-XamlWindow {
  param([Parameter(Mandatory)][string]$Xaml)
  $stringReader = New-Object System.IO.StringReader $Xaml
  $xmlReader = [System.Xml.XmlReader]::Create($stringReader)
  try {
    return [Windows.Markup.XamlReader]::Load($xmlReader)
  } finally {
    $xmlReader.Close()
    $stringReader.Dispose()
  }
}

function Show-OverrideDialog {
  param(
    [Parameter()] [System.Windows.Window] $Owner,
    [Parameter()] [pscustomobject] $Existing
  )

  $dialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Override Settings"
        SizeToContent="WidthAndHeight"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="#FFF5F7FB"
        Foreground="#FF1B2430">
  <Window.Resources>
    <SolidColorBrush x:Key="BackgroundBrush" Color="#FFF5F7FB"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFFFF"/>
    <SolidColorBrush x:Key="ForegroundBrush" Color="#FF1B2430"/>
    <SolidColorBrush x:Key="BorderBrush" Color="#FFD4DCEC"/>
    <SolidColorBrush x:Key="SecondaryTextBrush" Color="#FF5C6B84"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#FF2F6BFF"/>
    <SolidColorBrush x:Key="ListItemHoverBrush" Color="#FFE6EEFF"/>
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
    </Style>
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Border"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4">
              <ContentPresenter HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                Margin="4,2,4,2"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="Background" Value="{StaticResource ListItemHoverBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#FFD4DCEC"/>
                <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Border" Property="Opacity" Value="0.6"/>
                <Setter Property="Foreground" Value="{StaticResource SecondaryTextBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CaretBrush" Value="#FF1B2430"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="ComboBox">
      <Style.Resources>
        <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#FFFFFFFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#FFFFFFFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#FF2F6BFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#FFFFFFFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.ControlLightBrushKey}" Color="#FFE6EEFF"/>
        <SolidColorBrush x:Key="{x:Static SystemColors.ControlDarkBrushKey}" Color="#FFD4DCEC"/>
      </Style.Resources>
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="4,2"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="Padding" Value="6,3"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Style.Triggers>
        <Trigger Property="IsHighlighted" Value="True">
          <Setter Property="Background" Value="{StaticResource ListItemHoverBrush}"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
          <Setter Property="Foreground" Value="#FF252525"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid Margin="12" Background="{StaticResource BackgroundBrush}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="350"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <Label Content="Executable:" Grid.Row="0" Grid.Column="0" Margin="0,0,8,6" VerticalAlignment="Center"/>

    <!-- TextBox and Browse button share the same cell -->
    <Grid Grid.Row="0" Grid.Column="1" Margin="0,0,8,6">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <TextBox x:Name="ExeText" Grid.Column="0" MinWidth="240" Margin="0,0,6,0" VerticalAlignment="Center"/>
      <Button x:Name="BrowseButton" Grid.Column="1" Content="Browse" Padding="12,4" VerticalAlignment="Center"/>
    </Grid>

    <Label Content="CPU priority:" Grid.Row="2" Grid.Column="0" Margin="0,0,8,6" VerticalAlignment="Center"/>
    <ComboBox x:Name="CpuCombo" Grid.Row="2" Grid.Column="1" Margin="0,0,8,6" DisplayMemberPath="Name" SelectedValuePath="Value"/>

    <Label Content="I/O priority:" Grid.Row="3" Grid.Column="0" Margin="0,0,8,6" VerticalAlignment="Center"/>
    <ComboBox x:Name="IoCombo" Grid.Row="3" Grid.Column="1" Margin="0,0,8,6" DisplayMemberPath="Name" SelectedValuePath="Value"/>

    <Label Content="Memory page priority:" Grid.Row="4" Grid.Column="0" Margin="0,0,8,6" VerticalAlignment="Center"/>
    <ComboBox x:Name="PageCombo" Grid.Row="4" Grid.Column="1" Margin="0,0,8,6" DisplayMemberPath="Name" SelectedValuePath="Value"/>

    <StackPanel Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,12,8">
      <Button x:Name="SaveButton" Content="Save" Width="90" Margin="0,0,12,0"/>
      <Button x:Name="CancelButton" Content="Cancel" Width="90" Margin="0,0,0,0"/>
    </StackPanel>
  </Grid>
</Window>
'@

  $dialog = Load-XamlWindow -Xaml $dialogXaml
  if ($Owner) { $dialog.Owner = $Owner }

  $exeText    = $dialog.FindName('ExeText')
  $browseBtn  = $dialog.FindName('BrowseButton')
  $cpuCombo   = $dialog.FindName('CpuCombo')
  $ioCombo    = $dialog.FindName('IoCombo')
  $pageCombo  = $dialog.FindName('PageCombo')
  $saveBtn    = $dialog.FindName('SaveButton')
  $cancelBtn  = $dialog.FindName('CancelButton')

  $cpuCombo.ItemsSource = $CpuOptions
  $ioCombo.ItemsSource  = $IoOptions
  $pageCombo.ItemsSource= $PageOptions

  if ($Existing) {
    $exeText.Text = $Existing.Executable
    $cpuCombo.SelectedValue = $Existing.CpuValue
    if ($Existing.IoValue   -ne $null) { $ioCombo.SelectedValue   = $Existing.IoValue }   else { $ioCombo.SelectedIndex = 0 }
    if ($Existing.PageValue -ne $null) { $pageCombo.SelectedValue = $Existing.PageValue } else { $pageCombo.SelectedIndex = 0 }
  } else {
    $cpuCombo.SelectedValue = 2  # Normal
    $ioCombo.SelectedIndex  = 0  # (N/A)
    $pageCombo.SelectedIndex= 0  # (N/A)
  }

  $browseBtn.Add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Filter = 'Executable files (*.exe)|*.exe|All files (*.*)|*.*'
    $ofd.Multiselect = $false
    if ($ofd.ShowDialog()) {
      $exeText.Text = [System.IO.Path]::GetFileName($ofd.FileName)
    }
  })

  $dialog.Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq 'Enter') { $saveBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
    if ($e.Key -eq 'Escape') { $cancelBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
  })

  $cancelBtn.Add_Click({
    $dialog.DialogResult = $false
    $dialog.Close()
  })

  $saveBtn.Add_Click({
    $exeRaw = $exeText.Text
    if ([string]::IsNullOrWhiteSpace($exeRaw)) {
      [void][System.Windows.MessageBox]::Show('Executable name is required.','Validation', 'OK', 'Warning')
      return
    }
    $exeTrim = $exeRaw.Trim()
    if (-not $exeTrim.EndsWith('.exe', [System.StringComparison]::OrdinalIgnoreCase)) { $exeTrim = "$exeTrim.exe" }
    $exeName = [System.IO.Path]::GetFileName($exeTrim)
    if (-not $exeName) {
      [void][System.Windows.MessageBox]::Show('Executable name is invalid.','Validation', 'OK', 'Warning')
      return
    }

    $cpuSelected = $cpuCombo.SelectedValue
    if ($cpuSelected -eq $null) {
      [void][System.Windows.MessageBox]::Show('Choose a CPU priority level.','Validation', 'OK', 'Warning')
      return
    }

    $ioSelected   = $ioCombo.SelectedValue
    $pageSelected = $pageCombo.SelectedValue

    $result = [pscustomobject]@{
      Executable = $exeName
      CpuValue   = [int]$cpuSelected
      IoValue    = if ($ioSelected -eq $null -or $ioSelected -eq '') { $null } else { [int]$ioSelected }
      PageValue  = if ($pageSelected -eq $null -or $pageSelected -eq '') { $null } else { [int]$pageSelected }
    }

    $dialog.Tag = $result
    $dialog.DialogResult = $true
    $dialog.Close()
  })

  [void]$dialog.ShowDialog()
  if ($dialog.DialogResult -eq $true) {
    Write-Log "Dialog saved for executable '$($dialog.Tag.Executable)'" 'INFO'
    return $dialog.Tag
  }
  Write-Log 'Dialog cancelled without changes' 'INFO'
  return $null
}

function Launch-MainWindow {
  $mainWindowXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Process Priority Helper"
        Height="500"
        Width="560"
        MinHeight="400"
        MinWidth="560"
        WindowStartupLocation="CenterScreen"
        Background="#FFF5F7FB"
        Foreground="#FF1B2430">
  <Window.Resources>
    <SolidColorBrush x:Key="BackgroundBrush" Color="#FFF5F7FB"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFFFF"/>
    <SolidColorBrush x:Key="ForegroundBrush" Color="#FF1B2430"/>
    <SolidColorBrush x:Key="BorderBrush" Color="#FFD4DCEC"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#FF2F6BFF"/>
    <SolidColorBrush x:Key="SecondaryTextBrush" Color="#FF5C6B84"/>
    <SolidColorBrush x:Key="GridLineBrush" Color="#FFE0E4EE"/>
    <SolidColorBrush x:Key="RowBackgroundBrush" Color="#FFFFFFFF"/>
    <SolidColorBrush x:Key="RowAltBrush" Color="#FFF0F3FA"/>
    <SolidColorBrush x:Key="RowSelectedBrush" Color="#FF2F6BFF"/>
    <SolidColorBrush x:Key="RowSelectedForegroundBrush" Color="#FFFFFFFF"/>

    <Style TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Border"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="4">
              <ContentPresenter HorizontalAlignment="Center"
                                VerticalAlignment="Center"
                                Margin="4,2,4,2"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#FFF0F3FA"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Border" Property="Background" Value="#FFD4DCEC"/>
                <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Border" Property="Opacity" Value="0.6"/>
                <Setter Property="Foreground" Value="{StaticResource SecondaryTextBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="RowBackground" Value="{StaticResource RowBackgroundBrush}"/>
      <Setter Property="AlternatingRowBackground" Value="{StaticResource RowAltBrush}"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource GridLineBrush}"/>
      <Setter Property="VerticalGridLinesBrush" Value="{StaticResource GridLineBrush}"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
    </Style>

    <Style x:Key="CenterCell" TargetType="TextBlock">
      <Setter Property="HorizontalAlignment" Value="Center"/>
      <Setter Property="TextAlignment" Value="Center"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0"/>
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#FFF0F3FA"/>
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>

    <Style TargetType="DataGridRow">
      <Setter Property="Background" Value="{StaticResource RowBackgroundBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="SnapsToDevicePixels" Value="True"/>
      <Setter Property="MinHeight" Value="25"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{StaticResource RowSelectedBrush}"/>
          <Setter Property="Foreground" Value="{StaticResource RowSelectedForegroundBrush}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="DataGridCell">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource GridLineBrush}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Padding" Value="4,3"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="ClipToBounds" Value="True"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="{StaticResource RowSelectedBrush}"/>
          <Setter Property="Foreground" Value="{StaticResource RowSelectedForegroundBrush}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <DockPanel Margin="10">
    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,10">
      <Button x:Name="AddButton" Content="Add" Width="90" Margin="0,0,8,0"/>
      <Button x:Name="EditButton" Content="Edit" Width="90" Margin="0,0,8,0"/>
      <Button x:Name="RemoveButton" Content="Remove" Width="90"/>
    </StackPanel>

    <DataGrid x:Name="OverridesGrid"
              AutoGenerateColumns="False"
              CanUserAddRows="False"
              IsReadOnly="True"
              SelectionMode="Single"
              HeadersVisibility="Column"
              CanUserSortColumns="False"
              RowDetailsVisibilityMode="Collapsed"
              ScrollViewer.HorizontalScrollBarVisibility="Disabled"
              ScrollViewer.VerticalScrollBarVisibility="Auto"
              ScrollViewer.CanContentScroll="True"
              EnableRowVirtualization="True"
              EnableColumnVirtualization="True"
              VirtualizingPanel.IsVirtualizing="True"
              VirtualizingPanel.VirtualizationMode="Recycling">
      <DataGrid.Columns>

        <DataGridTemplateColumn Header="Executable" Width="*" MinWidth="160">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <Grid ClipToBounds="True">
                <TextBlock Text="{Binding Executable}"
                           TextWrapping="NoWrap"
                           TextTrimming="CharacterEllipsis"
                           VerticalAlignment="Center"
                           HorizontalAlignment="Center"
                           TextAlignment="Center"/>
                <TextBlock.ToolTip>
                  <ToolTip Placement="Mouse">
                    <TextBlock Text="{Binding Executable}" TextWrapping="Wrap" MaxWidth="400"/>
                  </ToolTip>
                </TextBlock.ToolTip>
              </Grid>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>

        <DataGridTextColumn Header="CPU Priority" Binding="{Binding CpuDisplay}" Width="120" ElementStyle="{StaticResource CenterCell}"/>
        <DataGridTextColumn Header="I/O Priority" Binding="{Binding IoDisplay}" Width="120" ElementStyle="{StaticResource CenterCell}"/>
        <DataGridTextColumn Header="Memory Priority" Binding="{Binding PageDisplay}" Width="140" ElementStyle="{StaticResource CenterCell}"/>

      </DataGrid.Columns>
    </DataGrid>
  </DockPanel>
</Window>
'@

  $window    = Load-XamlWindow -Xaml $mainWindowXaml
  $grid      = $window.FindName('OverridesGrid')
  $addBtn    = $window.FindName('AddButton')
  $editBtn   = $window.FindName('EditButton')
  $removeBtn = $window.FindName('RemoveButton')

  $refresh = {
    try {
      Write-Log 'Refreshing overrides from registry' 'INFO'
      $data = @(Get-Overrides)
      $grid.ItemsSource = $null
      $grid.ItemsSource = $data
      Write-Log ("Loaded {0} override(s)" -f $data.Count) 'INFO'
    } catch {
      Show-ErrorMessage "Failed to read overrides: $_"
      Write-Log "Refresh failed: $_" 'ERROR'
    }
  }

  $addBtn.Add_Click({
    try {
      Write-Log 'Opening Add dialog' 'INFO'
      $result = Show-OverrideDialog -Owner $window
      if ($result) {
        if ($result.CpuValue -eq 4) {
          $confirm = [System.Windows.MessageBox]::Show('Setting CPU Priority to Realtime can make the system unresponsive. Continue?', 'Warning', 'YesNo', 'Warning')
          if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }
        Apply-Override -Data $result
        $refresh.Invoke()
      }
    } catch {
      Show-ErrorMessage "Failed to apply override: $_"
      Write-Log "Add operation failed: $_" 'ERROR'
    }
  })

  $editHandler = {
    if (-not $grid.SelectedItem) {
      Show-InfoMessage 'Select an override first.'
      Write-Log 'Edit requested without a selection' 'WARN'
      return
    }
    $selected = $grid.SelectedItem
    try {
      Write-Log "Opening Edit dialog for '$($selected.Executable)'" 'INFO'
      $result = Show-OverrideDialog -Owner $window -Existing $selected
      if ($result) {
        if ($result.CpuValue -eq 4) {
          $confirm = [System.Windows.MessageBox]::Show('Setting Realtime priority can make the system unresponsive. Continue?', 'Warning', 'YesNo', 'Warning')
          if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }
        $oldName = $selected.Executable
        $newName = $result.Executable
        if ($oldName -and $newName -and (-not $newName.Equals($oldName, [System.StringComparison]::OrdinalIgnoreCase))) {
          try {
            if (Remove-PerfOptionsKey -Executable $oldName) {
              Write-Log "Removed previous override entry for '$oldName' due to rename" 'INFO'
            }
          } catch {
            Write-Log "Failed to remove previous override for '$oldName': $_" 'WARN'
          }
        }
        Apply-Override -Data $result
        $refresh.Invoke()
      }
    } catch {
      Show-ErrorMessage "Failed to update override: $_"
      Write-Log "Edit operation failed: $_" 'ERROR'
    }
  }

  # Only the Edit button can trigger editing
  $editBtn.Add_Click($editHandler)

  # Block double-clicks on the grid entirely
  $grid.Add_PreviewMouseDoubleClick({ param($s,$e) $e.Handled = $true })

  $removeBtn.Add_Click({
    if (-not $grid.SelectedItem) {
      Show-InfoMessage 'Select an override first.'
      Write-Log 'Remove requested without a selection' 'WARN'
      return
    }
    $selected = $grid.SelectedItem
    try {
      Write-Log "Removing override for '$($selected.Executable)'" 'INFO'
      if (Remove-PerfOptionsKey -Executable $selected.Executable) {
        Write-Log '  IFEO entry removed (including PerfOptions)' 'INFO'
      } elseif (Remove-OverrideValues -Executable $selected.Executable) {
        Write-Log '  Priority values removed (CPU, IO, Page)' 'INFO'
      } else {
        Write-Log '  Nothing to remove or failed to remove override' 'WARN'
      }
      $refresh.Invoke()
    } catch {
      Show-ErrorMessage "Failed to remove override: $_"
      Write-Log "Remove operation failed: $_" 'ERROR'
    }
  })

  $window.Add_SourceInitialized({ $refresh.Invoke() })
  [void]$window.ShowDialog()
}

try {
  Launch-MainWindow
} catch {
  Show-ErrorMessage "Fatal error: $_"
  Write-Log "Fatal error: $_" 'ERROR'
}