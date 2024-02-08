#Opens Powershell (hidden)
PowerShell.exe -WindowStyle hidden{

#Function for loading the XAML-Code (GUI Stuff)
    function Load-XAML {
    param (
        [string]$XAML
    )

    # Loading the Xaml
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $XAML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    $reader = (New-Object System.Xml.XmlNodeReader $XAML)
    return [Windows.Markup.XamlReader]::Load($reader)
}

# XAML-Code itself
$inputXML_MainWindow = @"
<Window x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp2"
        mc:Ignorable="d"
        Title="IP Scanner" SizeToContent="Height" Width="880" ResizeMode="CanMinimize" Background="#36393e" WindowStyle="SingleBorderWindow">
    <Grid>
        <!-- Adressbereich eingeben -->
        <Label Content="Specify address range:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" FontSize="16" Foreground="#FF0000"/>
        <TextBox Name="RangeTextBox" HorizontalAlignment="Left" Height="23" Margin="15,36,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="165" Foreground="#FF0000" Background="#303337" BorderBrush="Black" BorderThickness="1">
        <TextBox.Style>
            <Style TargetType="TextBox">
                <Setter Property="BorderBrush" Value="Black"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Style.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter Property="BorderBrush" Value="Black"/>
                    </Trigger>
                    <Trigger Property="IsFocused" Value="True">
                        <Setter Property="BorderBrush" Value="Black"/>
                    </Trigger>
                </Style.Triggers>
            </Style>
        </TextBox.Style>
        <TextBox.Template>
            <ControlTemplate TargetType="TextBox">
                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
                    <ScrollViewer x:Name="PART_ContentHost" />
                </Border>
            </ControlTemplate>
        </TextBox.Template>
        </TextBox>
    
        <!-- Hostnamen auflösen -->
        <Label Content="Resolve Hostname (SLOW!)" HorizontalAlignment="Left" Margin="29,72,0,0" VerticalAlignment="Top" FontSize="16" Foreground="#FF0000"/>
        <CheckBox Name="HostNameButton" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15,82,0,0" FontSize="16">
            <CheckBox.Style>
                <Style TargetType="CheckBox">
                    <Setter Property="Margin" Value="5"/>
                    <Setter Property="RenderTransformOrigin" Value="0.481,1.431"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="CheckBox">
                                <Grid>
                                    <Border x:Name="border" BorderBrush="Black" BorderThickness="1" Background="#303337"/>
                                    <Path x:Name="checkmark" Stroke="Red" StrokeThickness="2" Data="M2,8 L6,12 L14,2" Visibility="Collapsed"/>
                                </Grid>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsChecked" Value="True">
                                        <Setter TargetName="checkmark" Property="Visibility" Value="Visible"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter TargetName="border" Property="Background" Value="#36393e"/>
                                    </Trigger>
                                    <Trigger Property="IsPressed" Value="True">
                                        <Setter TargetName="border" Property="Background" Value="#303337"/>
                                    </Trigger>
                                    <Trigger Property="IsChecked" Value="False">
                                        <Setter TargetName="checkmark" Property="Visibility" Value="Hidden"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </CheckBox.Style>
        </CheckBox>
    
        <!-- Go-Button -->
        <Button Name="GoButton" Content="Go!" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15,115,0,0" Width="75" Height="30" FontSize="16" Foreground="#FF0000">
            <Button.Template>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="border" BorderBrush="Black" BorderThickness="1" Background="#303337">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#36393e"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Button.Template>
        </Button>

        <!-- Datagrid -->
        <DataGrid Name="Datagrid" AutoGenerateColumns="False" HeadersVisibility="Column" FontSize="16" Background="#333333" Foreground="#FF0000"
                  RowBackground="#36393e" AlternatingRowBackground="#36393e" HorizontalScrollBarVisibility="Hidden" 
                  IsReadOnly="True" CanUserSortColumns="False" CanUserReorderColumns="False" SelectionUnit="Cell" Margin="0,150,0,0">
                <DataGrid.Resources>
                  <Style TargetType="{x:Type DataGridColumnHeader}">
                      <Setter Property="Background" Value="#282b30"></Setter> <!-- Sehr dunkles Grau für Überschrift -->
                      <Setter Property="Foreground" Value="#FF0000"></Setter> <!-- Rote Textfarbe für Überschrift -->
                      <Setter Property="FontWeight" Value="SemiBold"></Setter>
                  </Style>
                  <Style TargetType="{x:Type DataGridCell}">
                      <Setter Property="Background" Value="#36393e"></Setter> <!-- Minimal helleres Grau für Zellen -->
                      <Setter Property="BorderBrush" Value="#282b30"></Setter> <!-- Passende Farbe für Zellenrahmen -->
                      <Setter Property="BorderThickness" Value="0.9"></Setter> <!-- Optional: Dicke des Zellenrahmens -->
                  </Style>
                  <Style TargetType="{x:Type DataGridRow}">
                      <Setter Property="BorderBrush" Value="#282b30"></Setter> <!-- Passende Farbe für Rahmen der gesamten Zeile -->
                      <Setter Property="BorderThickness" Value="0.5"></Setter> <!-- Optional: Dicke des Rahmens der gesamten Zeile -->
                  </Style>
                </DataGrid.Resources>
              
              <DataGrid.Style>
                  <Style TargetType="{x:Type DataGrid}">
                      <Setter Property="BorderBrush" Value="#282b30"></Setter> <!-- Passende Farbe für Rahmen der gesamten DataGrid -->
                      <Setter Property="BorderThickness" Value="0.5"></Setter> <!-- Optional: Dicke des Rahmens der gesamten DataGrid -->
                  </Style>
              </DataGrid.Style>
            
            <DataGrid.Columns>
                <!-- Datagrid-Spalten -->
                <DataGridTextColumn Header="IP" Binding="{Binding IP}" Width="Auto" MinWidth="100"/>
                <DataGridTextColumn Header="MAC" Binding="{Binding MAC}" Width="Auto" MinWidth="100"/>
                <DataGridTextColumn Header="Hostname" Binding="{Binding Hostname}" Width="Auto" MinWidth="100"/>
                <DataGridTextColumn Header="Vendor" Binding="{Binding Hersteller}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
</Window>

"@

$MainWindow = Load-XAML -XAML $inputXML_MainWindow

$HostNameCheckBox = $MainWindow.FindName("HostNameButton")

$GoButton = $MainWindow.FindName("GoButton")

$RangeTextBox = $MainWindow.FindName("RangeTextBox")

$DataGrid = $MainWindow.FindName("Datagrid")

$DataGrid.Visibility= "Hidden"
############################################################

#Imports the CSV (Vendorlist) from my github
if(!(Test-Path C:\temp\mac-vendors-export.csv)){
    $GitRequest = Invoke-WebRequest https://raw.githubusercontent.com/Neo0412/IPScan/main/mac-vendors-export.csv
    $CSVasText = $GitRequest.Content
    $CSVasText |Out-File -FilePath C:\temp\mac-vendors-export.csv
}

$MACRecords=Import-CSV "C:\temp\mac-vendors-export.csv"

$GoButton.Add_Click({

    #Formatting the user given IP
    $FullStartIP,$EndIP = $RangeTextBox.Text -split '-'
    $StartIP = [int]($FullStartIP -split '\.')[3]
    $IPWithoutLastByte = ($FullStartIP -split '\.')[0..2] -join '.'

    #Creates an array with every IP of the IP range
    $IPRange = $StartIP..$EndIP | ForEach-Object { "$IPWithoutLastByte"+'.'+"$_" }

    #Formatting the array so that is usable for the following command
    $query = $IPRange -join "' or Address='"

    #Pings are executed
    $IPs = Get-CimInstance -ClassName Win32_PingStatus -Filter "(Address='$query') and timeout=1000"
    $DevicesOnly = $IPs | Where-Object{$_.StatusCode -eq 0}

    $Data = @()

    #Loop for gathering infos about each IP
    foreach($Device in $DevicesOnly){

        $IPAddress = $Device.ProtocolAddress
        
        #Gets MAC address
        $MAC = arp -a $IPAddress | Select-String '([0-9a-f]{2}-){5}[0-9a-f]{2}' | Select-Object -Expand Matches | Select-Object -Expand Value

        if($HostNameCheckBox.IsChecked -eq $true){
        
            #DNS name is resolved
            $DNSName = Resolve-DNSName -Name $IPAddress -Server 8.8.8.8 -QuickTimeout -ErrorAction SilentlyContinue| Select-Object NameHost
            $HostName = $DNSName.NameHost
        }

        $Vendor = $null

        #Vendor is determit threw the MAC address
        if ($MAC -ne $null){

            $3LetterMAC = (($MAC.SubString(0,8)).ToUpper()) -replace '-', ':'

            ForEach ($Record in $MACRecords){ 

                If ($Record.'MAC Prefix' -eq $3LetterMAC) {
                $Vendor = $Record.'Vendor Name'   
                break     
                }
                
            }
        }

        #Creating a PS-CustomObject
        $Temp = New-Object -TypeName PSCustomObject
        $Temp | Add-Member -Type NoteProperty -Name IP -Value $IPAddress
        $Temp | Add-Member -Type NoteProperty -Name MAC -Value $MAC
        $Temp | Add-Member -Type NoteProperty -Name "Hostname" -Value $HostName
        $Temp | Add-Member -Type NoteProperty -Name "Hersteller" -Value $Vendor

        $Data += $Temp
    }

    # Adding the data to the GUI datagrid
    $MainWindow.FindName("Datagrid").ItemsSource = $Data
    $DataGrid.Visibility= "Visible"
})


# Display GUI
$MainWindow.ShowDialog()

if(Test-Path C:\temp\mac-vendors-export.csv){
    Remove-Item C:\temp\mac-vendors-export.csv -Force
}
}