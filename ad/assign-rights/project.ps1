# Enable WPF
Add-Type -AssemblyName PresentationCore,PresentationFramework

# One table to rule them all...
$global:syncHash = [HashTable]::Synchronized(@{})

# Get all groups in the forest
Function Get-AllGroups {

    Begin {
        $results = [System.Collections.ArrayList]::new()

        $searcher = [System.DirectoryServices.DirectorySearcher]::new()
        $searcher.Filter = "(&(objectCategory=group)(managedBy=*))"
        $searcher.PropertiesToLoad.Add("Name") | Out-Null
        $searcher.PropertiesToLoad.Add("Description") | Out-Null
        $searcher.PropertiesToLoad.Add("DistinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("ManagedBy") | Out-Null
        $searcher.PropertiesToLoad.Add("Member") | Out-Null
    }
    
    Process {
        $allGroups = $searcher.FindAll()

        For ($i=0; $i -lt $allGroups.Count; $i++) {
            $entry = ($allGroups[$i]).Properties

            $entryInformation = [PSCustomObject]@{
                Name                = [string]$entry.name
                Details             = [string]$entry.description
                DistinguishedName   = [string]$entry.distinguishedname
                ManagedBy           = [string]$entry.managedby
                Members             = [string]$entry.member
            }

            $results.Add($entryInformation) | Out-Null
        }

    }

    End {
        $searcher.Dispose()
        Return $results
    }
    
}

# Get all users in AD that aren't disabled
Function Get-AllUsers {

    Begin {
        $results = [System.Collections.ArrayList]::new()

        $searcher = [System.DirectoryServices.DirectorySearcher]::new()
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
        $searcher.PropertiesToLoad.Add("Name") | Out-Null
        $searcher.PropertiesToLoad.Add("DistinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null
    }

    Process {
        $allUsers = $searcher.FindAll()
        
        For ($i=0; $i -lt $allUsers.Count; $i++) {
            $entry = ($allUsers[$i]).Properties

            $entryInformation = [PSCustomObject]@{
                Name = [string]$entry.name
                DN   = [string]$entry.distinguishedname
                SAM  = [string]$entry.samaccountname
            }

            $results.Add($entryInformation) | Out-Null
        }
    }

    End {
        $searcher.Dispose()
        Return $results
    }

}

# Get groups that a user belongs to
Function Get-UserGroups {
    Param(
        $User
    )

    Begin {
        $searcher = [System.DirectoryServices.DirectorySearcher]::new()
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$User))"
        $searcher.PropertiesToLoad.Add("MemberOf") | Out-Null
    }

    Process {
        $userGroups = ($searcher.FindAll()).Properties
        $results = [PSCustomObject]@{
            User   = $User
            Groups = $userGroups.memberof
        }
    }

    End {
        $searcher.Dispose()
        Return $results
    }

}

Function Set-ProgressMessage {

    Param(
        [String]$message
    )

    Do {
        Start-Sleep -Milliseconds 10
    } Until ($syncHash.InfoBox.Dispatcher -ne $null)

    If ($message){
        $syncHash.InfoBox.Dispatcher.Invoke(
            [action]{$syncHash.InfoBox.Text = $message},"Normal"
        )
    }

}

Function Show-ProgressWindow {

    # Create a new runspace for the boxes to run in
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

    # Build the xml as an array to allow variables
    $command = [PowerShell]::Create().AddScript({

        [Xml]$xml = '
            <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Add/Remove Users"
                Height="250" Width="270" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
                <Grid>
                    <StackPanel>
                        <TextBlock Name="Text1" HorizontalAlignment="Center" Margin="0,75,0,5" Text="Loading. Please wait."/>
                        <ProgressBar x:Name="PBar" HorizontalAlignment="Center" Width="45" Height="20" IsIndeterminate="True"/>
                        <TextBlock Name="Text2" HorizontalAlignment="Center" Margin="0,70,0,5" Text=""/>
                    </StackPanel>
                </Grid>
            </Window>
        '
         
        # Window Constructor
        $reader = New-Object System.Xml.XmlNodeReader $xml
        $syncHash.ProgressWindow = [Windows.Markup.XamlReader]::Load($reader)

        # Object identification
        $syncHash.AutoClose = $true
        $syncHash.InfoBox = $syncHash.ProgressWindow.FindName("Text2")

        # Handle the 'X' button
        $syncHash.ProgressWindow.Add_Closing({
            If ($syncHash.AutoClose -ne $true) {
                $command.EndInvoke($result)
                $command.Runspace.Dispose()
                $command.Runspace.Close()
                $command.Dispose()
                Break
            }
        })

        # Show the window to the user
        [Void]$syncHash.ProgressWindow.ShowDialog()
        $command.EndInvoke($result)
        $command.Runspace.Dispose()
        $command.Runspace.Close()
        $command.Dispose()
    })

    # Create tracking then open the runspace
    $command.Runspace = $runspace
    $result = $command.BeginInvoke()

}

Function Close-Window {
    $syncHash.ProgressWindow.Dispatcher.Invoke(
        [action]{$syncHash.ProgressWindow.Close()},"Normal"
    )
}

Function Show-MainWindow {

    Function Show-Details {
        $selected = $this
        $i = $selected.Name -Replace "Group"
        $details = ($managedGroups[$i]).Details

        $syncHash.MainWindow.FindName("Details").Dispatcher.Invoke(
            [action]{$syncHash.MainWindow.FindName("Details").Text = $details},"Normal"
        )

    }

    Function Show-GroupMembers {
        $selected = $this
        $i = $selected.Name -Replace "Group"
        $members = ($managedGroups[$i]).Members
        $selected = $syncHash.MainWindow.FindName("Groups").SelectedItems.Count

        If ($selected -gt 1) {
            $available = $synchHash.MainWindow.FindName("AddRemove").Items
            Write-Host $available
            Write-Host $available.Count
        }
        Else {
            If ($members) {
                $members = $members | Sort-Object
                For ($i=0;$i -lt $members.Count;$i++) {
                    $newListItem = New-Object System.Windows.Controls.ListBoxItem
                    $newListItem.Name     = "Member$i"
                    $newListItem.Content  = $members[$i]

                    [void]$syncHash.MainWindow.FindName("AddRemove").Items.Add($newListItem)
                    [void]$syncHash.MainWindow.RegisterName($newListItem.Name, $newListItem)
                }
            }
        }

    }

    Function Clear-DetailsAndUsers {
        $status = $syncHash.MainWindow.FindName("Groups").SelectedIndex

        If ($status -eq -1) {
            $items = $syncHash.MainWindow.FindName("AddRemove").Items | Select-Object *

            ForEach ($item in $items) {
                $syncHash.MainWindow.UnregisterName($item.Name)
            }

            $syncHash.MainWindow.FindName("Details").Dispatcher.Invoke(
                [action]{$syncHash.MainWindow.FindName("Details").Text = $null},"Normal"
            )

            $syncHash.MainWindow.FindName("AddRemove").Dispatcher.Invoke(
                [action]{$syncHash.MainWindow.FindName("AddRemove").Items.Clear()},"Normal"
            )
        }
    }
    # define the main window
    [xml]$code = '
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        x:Name="Window" Title="Add/Remove Users" Height="680" Width="555" ShowInTaskbar="True"
        MinHeight="355" MinWidth="555" WindowStartupLocation="CenterScreen">

            <Window.Resources>
                <Style x:Key="ButtonStyle" TargetType="Button">
                    <Setter Property="Height" Value="45"/>
                    <Setter Property="Width" Value="70"/>
                    <Setter Property="FontSize" Value="16"/>
                    <Setter Property="FontWeight" Value="Bold"/>
                    <Setter Property="HorizontalContentAlignment" Value="Center"/>
                    <Setter Property="Margin" Value="2"/>
                </Style>
            </Window.Resources>

            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="35"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="75"/>
                </Grid.RowDefinitions>

                <!-- First Row -->
                <Grid Grid.Row="0">
                    <TextBlock x:Name="LoggedIn" HorizontalAlignment="Right" VerticalAlignment="Top"
                        Margin="0,10,10,0" Width="150" Height="25" Text="Logged in as:"/>
                </Grid>

                <!-- Second Row -->
                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <Grid Grid.Column="0">
                        <ListBox x:Name="Groups" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                            Margin="10,5,5,5" SelectionMode="Multiple"/>
                    </Grid>

                    <Grid Grid.Column="1">
                        <TextBox x:Name="Details" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                            Margin="5,5,10,5" IsReadOnly="True" TextWrapping="Wrap"/>
                    </Grid>
                </Grid>

                <!-- Third Row -->
                <Grid Grid.Row="2">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="75"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <Grid Grid.Column="0">
                        <ListBox x:Name="Users" VerticalAlignment="Stretch" HorizontalAlignment="Stretch"
                            Margin="10,5,5,5"/>
                    </Grid> 

                    <Grid Grid.Column="1">
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                            <Button x:Name="MoveTo" Content="&#x1f846;" Style="{StaticResource ButtonStyle}"/>
                            <Button x:Name="MoveFrom" Content="&#x1f844;" Style="{StaticResource ButtonStyle}"/>
                        </StackPanel>
                    </Grid>

                    <Grid Grid.Column="2">
                    <ListBox x:Name="AddRemove" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                        Margin="5,5,10,5"/>
                    </Grid>
                </Grid>

                <!-- Fourth Row -->
                <Grid Grid.Row="3">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="75"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <Grid Grid.Column="0">
                        <TextBox x:Name="Filter" HorizontalAlignment="Stretch" Height="23" Margin="10,5,5,5"/>
                    </Grid>
                    
                    <Grid Grid.Column="1">
                        <TextBlock x:Name="Spacer" Margin="2"/>
                    </Grid>

                    <Grid Grid.Column="2">
                        <Button x:Name="Apply" Content="Apply" Style="{StaticResource ButtonStyle}"
                            HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Grid>
                </Grid>
            </Grid>
        </Window>
    '
    $reader = New-Object System.Xml.XmlNodeReader $code
    $syncHash.MainWindow = [Windows.Markup.XamlReader]::Load($reader)
           
    # Add Groups to the Group element
    Set-ProgressMessage -message "Loading groups to window"
    For ($i=0;$i -lt $managedGroups.Count;$i++) {
        $group = $managedGroups.Name[$i]

        $newListItem = New-Object System.Windows.Controls.ListBoxItem
        $newListItem.Name       = "Group$i"
        $newListItem.Content    = $group

        $newListItem.Add_Selected({ Show-Details; Show-GroupMembers })
        [void]$syncHash.MainWindow.FindName("Groups").Items.Add($newListItem)
        [void]$syncHash.MainWindow.RegisterName($newListItem.Name, $newListItem)
    }

    # Add Users to the User element
    Set-ProgressMessage -message "Loading users to window"
    For ($i=0;$i -lt $allUsers.Count;$i++) {
        $user = $allUsers.Name[$i]

        $newListItem = New-Object System.Windows.Controls.ListBoxItem
        $newListItem.Name       = "User$i"
        $newListItem.Content    = $user

        [void]$syncHash.MainWindow.FindName("Users").Items.Add($newListItem)
        [void]$syncHash.MainWindow.RegisterName($newListItem.Name, $newListItem)
    }

    # Add a listener to the Listbox
    $syncHash.MainWindow.FindName("Groups").Add_SelectionChanged({ Clear-DetailsAndUsers })

    Close-Window
    [void]$syncHash.MainWindow.ShowDialog()
}

# ----------------------------------------------------------------------------------------------------------- #
# Start application; show user a progress window
Show-ProgressWindow
Set-ProgressMessage -message "Gathering group information"

# Get bulk information from AD
$allGroups = Get-AllGroups
$Script:allUsers = Get-AllUsers | Sort-Object Name
$myGroups = Get-UserGroups -User $env:username

# Filter info down and remove extraneous
$filteredGroups = [System.Collections.ArrayList]::new()
Set-ProgressMessage -message "Filtering groups"

ForEach ($group in $allGroups) {
    If ($group.managedby -in $myGroups.Groups) {
        $filteredGroups.Add($group) | Out-Null
    }
}

$managedGroups = [System.Collections.Generic.List[Object]]::new()
$regex = [Regex] '(?= CN=)'

For ($i=0;$i -lt $filteredGroups.Count;$i++) {
    $group = $filteredGroups[$i]
    $members = $group.Members

    If ($members) {
        $split = $regex.Split($members).Trim()
        $members = $split | ForEach-Object {($_ -Split ",")[0] -Replace "CN="}
    }

    $add = [PSCustomObject]@{
        Name                = $group.Name
        Details             = $group.Details
        DistinguishedName   = $group.DistinguishedName
        ManagedBy           = $group.ManagedBy
        Members             = $members
    }

    $managedGroups.Add($add)
}

$allGroups = $null
$myGroups = $null
$Script:managedGroups = $managedGroups | Sort-Object Name

# Close progress window and open main form
Show-MainWindow