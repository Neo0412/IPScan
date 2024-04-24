# Function for loading the XAML code
function Load-XAML {
    
    param (
        [string]$XAML
    )

    # Loading the XAML file
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $XAML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    $reader = (New-Object System.Xml.XmlNodeReader $XAML)
    return [Windows.Markup.XamlReader]::Load($reader)
}

# XAML code for the GUI
$inputXML_MainWindow = @"
<Window x:Class="WpfApp2.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:local="clr-namespace:WpfApp2"
        Title="IP Scanner" SizeToContent="Height" Width="880" ResizeMode="CanMinimize" Background="#36393e" WindowStyle="None" BorderThickness="1">
    <Window.Resources>
        <!-- Definition of a template for the ContextMenu -->
        <ItemsPanelTemplate x:Key="MenuItemPanelTemplate">
            <StackPanel Margin="-1,0,0,0" Background="#36393e"/>
        </ItemsPanelTemplate>
        <!-- Style for the ContextMenu -->
        <Style TargetType="{x:Type ContextMenu}">
            <Setter Property="ItemsPanel" Value="{StaticResource MenuItemPanelTemplate}"/>
        </Style>
        <!-- Style for MenuItem -->
        <Style TargetType="{x:Type MenuItem}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#FF0000"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type MenuItem}">
                    <Border Background="{TemplateBinding Background}" BorderThickness="0" MinWidth="130" MinHeight="30">
                    <ContentPresenter Content="{TemplateBinding Header}" HorizontalAlignment="Left" VerticalAlignment="Center"/>
                </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid>
        <!-- Taskbar -->
        <Rectangle Fill="#282b30" HorizontalAlignment="Left" Height="31" Stroke="Black" VerticalAlignment="Top" Width="880"/>
        <TextBlock X:Name="Taskbar" Text="IP Scanner" Margin="5" FontSize="16" Foreground="#FF0000"/>
        <!-- Red circle button to close -->
        <Button Name="ExitButton" Width="25" Height="25" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="10,0,0,0" Background="Transparent" BorderBrush="Transparent">
            <Button.Template>
                <ControlTemplate TargetType="Button">
                    <Ellipse x:Name="RedEllipse" Width="15" Height="15" Fill="Red" Stroke="Black" StrokeThickness="1.5"/>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="RedEllipse" Property="Fill" Value="DarkRed"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Button.Template>
        </Button>
        <!-- Yellow circle button to minimize -->
        <Button Name="MinimizeButton" Width="25" Height="25" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="800,0,25,0" Background="Transparent" BorderBrush="Transparent">
            <Button.Template>
                <ControlTemplate TargetType="Button">
                    <Ellipse x:Name="YellowEllipse" Width="15" Height="15" Fill="Yellow" Stroke="Black" StrokeThickness="1.5"/>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="YellowEllipse" Property="Fill" Value="#8B8000"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Button.Template>
        </Button>
        <!-- Input field for IP range -->
        <Label Content="Enter IP range:" HorizontalAlignment="Left" Margin="10,36,0,0" VerticalAlignment="Top" FontSize="16" Foreground="#FF0000"/>
        <TextBox Name="RangeTextBox" HorizontalAlignment="Left" Height="23" Margin="15,73,0,0" VerticalAlignment="Top" Width="180" Foreground="#FF0000" Background="#303337" BorderBrush="Black" BorderThickness="1">
            <!-- Style for the input field -->
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
                        <ScrollViewer x:Name="PART_ContentHost"/>
                    </Border>
                </ControlTemplate>
            </TextBox.Template>
        </TextBox>
        <!-- Checkbox to resolve hostnames -->
        <Label Content="Resolve hostnames (SLOW!)" HorizontalAlignment="Left" Margin="29,101,0,0" VerticalAlignment="Top" FontSize="16" Foreground="#FF0000"/>
        <CheckBox Name="HostNameButton" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15,111,0,0" FontSize="16">
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
        <!-- Go button -->
        <Button Name="GoButton" Content="Go!" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="15,150,0,0" Width="75" Height="30" FontSize="16" Foreground="#FF0000">
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
        <!-- Data grid for results -->
        <DataGrid Name="Datagrid" AutoGenerateColumns="False" HeadersVisibility="Column" FontSize="16" Background="#333333" Foreground="#FF0000"
                  RowBackground="#36393e" AlternatingRowBackground="#36393e" HorizontalScrollBarVisibility="Hidden" 
                  IsReadOnly="True" CanUserSortColumns="False" CanUserReorderColumns="False" SelectionUnit="Cell" Margin="0,200,0,0">
            <DataGrid.Resources>
                <!-- Style for headers -->
                <Style TargetType="{x:Type DataGridColumnHeader}">
                    <Setter Property="Background" Value="#282b30"/>
                    <Setter Property="Foreground" Value="#FF0000"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="BorderThickness" Value="1"/>
                    <Setter Property="BorderBrush" Value="Black"/>
                </Style>
                <!-- Style for cells -->
                <Style TargetType="{x:Type DataGridCell}">
                    <Setter Property="Background" Value="#36393e"/>
                    <Setter Property="BorderBrush" Value="#282b30"/>
                    <Setter Property="BorderThickness" Value="0.9"/>
                </Style>
                <!-- Style for rows -->
                <Style TargetType="{x:Type DataGridRow}">
                    <Setter Property="BorderBrush" Value="#282b30"/>
                    <Setter Property="BorderThickness" Value="0.5"/>
                </Style>
            </DataGrid.Resources>
            <!-- Style for the entire data grid -->
            <DataGrid.Style>
                <Style TargetType="{x:Type DataGrid}">
                    <Setter Property="BorderBrush" Value="#282b30"/>
                    <Setter Property="BorderThickness" Value="0.5"/>
                </Style>
            </DataGrid.Style>
            <!-- Columns for results -->
            <DataGrid.Columns>
                <DataGridTextColumn Header="IP" Binding="{Binding IP}" Width="Auto" MinWidth="100">
                    <DataGridTextColumn.CellStyle>
                        <Style TargetType="DataGridCell">
                            <!-- Context menu for cells -->
                            <Setter Property="ContextMenu">
                                <Setter.Value>
                                    <ContextMenu Background="#36393e" Foreground="#FF0000" FontWeight="SemiBold" BorderBrush="Black" BorderThickness="1.5">
                                        <MenuItem Name="OpenBrowser" Header="Open in Browser" IsCheckable="False"/>
                                        <MenuItem Name="Ping" Header="Run Ping" IsCheckable="False"/>
                                    </ContextMenu>
                                </Setter.Value>
                            </Setter>
                            <Style.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="#36393e"/>
                                    <Setter Property="BorderThickness" Value="0"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </DataGridTextColumn.CellStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="MAC" Binding="{Binding MAC}" Width="Auto" MinWidth="100"/>
                <DataGridTextColumn Header="Hostname" Binding="{Binding Hostname}" Width="Auto" MinWidth="100"/>
                <DataGridTextColumn Header="Manufacturer" Binding="{Binding Manufacturer}" Width="*"/>
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

$ExitButton = $MainWindow.FindName("ExitButton")

$MinimizeButton = $MainWindow.FindName("MinimizeButton")

$Taskbar = $MainWindow.FindName("Taskbar")

$ContextMenu = $DataGrid.Columns.CellStyle.Setters.Value

$OpenBrowserButton = $ContextMenu.Items[0]

$StartPing = $ContextMenu.Items[1]


$DataGrid.Visibility= "Hidden"
############################################################

# Importing the CSV
if(!(Test-Path C:\temp\mac-vendors-export.csv)){
    $GitRequest = Invoke-WebRequest https://raw.githubusercontent.com/Neo0412/IPScan/main/mac-vendors-export.csv
    $CSVasText = $GitRequest.Content
    $CSVasText |Out-File -FilePath C:\temp\mac-vendors-export.csv
}

$MACRecords=Import-CSV "C:\temp\mac-vendors-export.csv"

$GoButton.Add_Click({

    $IPRange = @()

    $StartIP, $EndIP = $RangeTextBox.Text -split "-" 

    if (!($EndIP -match "\.\d+\.\d+\.\d+")){
        
        # Formatting the IP
        $StartIPLastByte = [int]($StartIP -split '\.')[3]
        $IPWithoutLastByte = ($StartIP -split '\.')[0..2] -join '.'
    
        # Array with all IP addresses in the range is created
        $IPRange = $StartIPLastByte..$EndIP | ForEach-Object { "$IPWithoutLastByte"+'.'+"$_" }
        
    }
    else{
    
    
        # Function to convert an IP address to a decimal number
        function ConvertToDecimal($IP) {
            $IPArray = $IP.Split('.')
            $Decimal = 0
            for ($i = 0; $i -lt 4; $i++) {
                $Decimal += [int]::Parse($IPArray[$i]) * [math]::Pow(256, 3 - $i)
            }
            return $Decimal
        }
    
        # Function to convert a decimal number to an IP address
        function ConvertToIP($Decimal) {
            $IP = ""
            for ($i = 3; $i -ge 0; $i--) {
                $IP += ($Decimal -shr ($i * 8)) -band 255
                if ($i -gt 0) {
                    $IP += "."
                }
            }
            return $IP
        }
    
        # Convert the start and end IP addresses to decimal numbers
        $StartDecimal = ConvertToDecimal $StartIP
        $EndDecimal = ConvertToDecimal $EndIP
    
        # Generate all IP addresses in the range
        for ($i = $StartDecimal; $i -le $EndDecimal; $i++) {
        
            $TempIP = ConvertToIP $i
            $IPRange += $TempIP
        
        }
    
    }

    # List made usable for Win32_PingStatus
    $query = $IPRange -join "' or Address='"

    # Ping is executed and reachable IPs are stored in $DevicesOnly
    $IPs = Get-CimInstance -ClassName Win32_PingStatus -Filter "(Address='$query') and timeout=1000"
    $DevicesOnly = $IPs | Where-Object{$_.StatusCode -eq 0}

    $Data = @()

    # Loop to get information about each IP
    foreach($Device in $DevicesOnly){

        $IPAddress = $Device.ProtocolAddress
        
        # MAC address is determined
        $MAC = arp -a $IPAddress | Select-String '([0-9a-f]{2}-){5}[0-9a-f]{2}' | Select-Object -Expand Matches | Select-Object -Expand Value

        if($HostNameCheckBox.IsChecked -eq $true){
        
            # DNS name is attempted to be resolved
            $DNSName = Resolve-DNSName -Name $IPAddress -Server 8.8.8.8 -QuickTimeout -ErrorAction SilentlyContinue| Select-Object NameHost
            $HostName = $DNSName.NameHost
        }

        $Vendor = $null

        # Manufacturer is determined using the MAC address
        if ($MAC -ne $null){

            $3LetterMAC = (($MAC.SubString(0,8)).ToUpper()) -replace '-', ':'

            ForEach ($Record in $MACRecords){ 

                If ($Record.'MAC Prefix' -eq $3LetterMAC) {
                $Vendor = $Record.'Vendor Name'   
                break     
                }
                
            }
        }

        # PSCustomObject is populated
        $Temp = New-Object -TypeName PSCustomObject
        $Temp | Add-Member -Type NoteProperty -Name IP -Value $IPAddress
        $Temp | Add-Member -Type NoteProperty -Name MAC -Value $MAC
        $Temp | Add-Member -Type NoteProperty -Name "Hostname" -Value $HostName
        $Temp | Add-Member -Type NoteProperty -Name "Manufacturer" -Value $Vendor

        $Data += $Temp
    }

    # Assigning the data to the data grid
    $MainWindow.FindName("Datagrid").ItemsSource = $Data
    $DataGrid.Visibility= "Visible"
})


# Exit Button gets functionality
$ExitButton.Add_Click({
    $MainWindow.Close()
    exit
})

# Minimize Button gets functionality
$MinimizeButton.Add_Click({
    $MainWindow.WindowState = "Minimized"
})

# Function to move the GUI window
$Taskbar.Add_MouseLeftButtonDown({
    $MainWindow.DragMove()
})

$OpenBrowserButton.Add_Click({

    if($DataGrid.SelectedCells.Count -gt 0 -and $DataGrid.SelectedCells[0].Column.DisplayIndex -eq 0){
        
        # Retrieving and displaying cell value
        $SelectedCell = $DataGrid.SelectedCells[0]
        $CellValue = $SelectedCell.Item.$($SelectedCell.Column.Header)

        Start-Process "https://$CellValue"
    }

})

$StartPing.Add_Click({

    if($DataGrid.SelectedCells.Count -gt 0 -and $DataGrid.SelectedCells[0].Column.DisplayIndex -eq 0){

        # Retrieving and displaying cell value
        $SelectedCell = $DataGrid.SelectedCells[0]
        $CellValue = $SelectedCell.Item.$($SelectedCell.Column.Header)

        Start-Process "Cmd.exe" "/C Ping $cellValue -t"

    }

})

# Displaying the form
$MainWindow.ShowDialog()

if(Test-Path C:\temp\mac-vendors-export.csv){
    Remove-Item C:\temp\mac-vendors-export.csv -Force
}
