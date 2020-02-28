# Enable WPF
Add-Type -AssemblyName PresentationCore,PresentationFramework

# Enable Account Management
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# One table to rule them all...
$global:syncHash = [HashTable]::Synchronized(@{})

# Define some table segments
$syncHash.FilterUsers = [System.Collections.ArrayList]::new()
$syncHash.AddRemove = [System.Collections.ArrayList]::new()
$syncHash.AddButton = [System.Collections.ArrayList]::new()
$syncHash.RemoveButton = [System.Collections.ArrayList]::new()

# Get all groups in the forest
Function Get-AllGroups {

    Begin {
        $results = [System.Collections.ArrayList]::new()

        $searcher = [System.DirectoryServices.DirectorySearcher]::new()
        $searcher.Filter = "(&(objectCategory=group)(managedBy=*))"
        $searcher.PropertiesToLoad.Add("Name") | Out-Null
        $searcher.PropertiesToLoad.Add("Description") | Out-Null
        $searcher.PropertiesToLoad.Add("DistinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null
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
                sAMAccountName      = [string]$entry.samaccountname
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
        $results = [System.Data.DataTable]::new()
        $results.TableName = "AllUsers"

        $searcher = [System.DirectoryServices.DirectorySearcher]::new()
        $searcher.PageSize = 1000
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
        $searcher.PropertiesToLoad.Add("Name") | Out-Null
        $searcher.PropertiesToLoad.Add("DistinguishedName") | Out-Null
        $searcher.PropertiesToLoad.Add("sAMAccountName") | Out-Null

        [void]$results.Columns.Add("Name")
        [void]$results.Columns.Add("DN")
        [void]$results.Columns.Add("SAM")
    }

    Process {
        $allUsers = $searcher.FindAll()
        
        For ($i=0; $i -lt $allUsers.Count; $i++) {
            $entry = ($allUsers[$i]).Properties

                [void]$results.Rows.Add(
                    [string]$entry.name,
                    [string]$entry.distinguishedname,
                    [string]$entry.samaccountname
                )
        }
    }

    End {
        $searcher.Dispose()
        ,$results
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
        $searcher.PropertiesToLoad.Add("DistinguishedName") | Out-Null
    }

    Process {
        $userGroups = ($searcher.FindAll()).Properties
        $results = [PSCustomObject]@{
            User   = $User
            Groups = $userGroups.memberof
            DN     = $userGroups.distinguishedname
        }
    }

    End {
        $searcher.Dispose()
        Return $results
    }

}

# Filter down groups and create a table segment
Function Get-ManagedGroups {
    $allGroups = Get-AllGroups
    $myGroups = Get-UserGroups -User $env:username
    
    # Filter info down and remove extraneous
    $filteredGroups = [System.Collections.ArrayList]::new()
    
    ForEach ($group in $allGroups) {
        If (($group.managedby -in $myGroups.Groups) -or
        ($group.managedby -eq $myGroups.DN)) {
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
            sAMAccountName      = $group.sAMAccountName
            ManagedBy           = $group.ManagedBy
            Members             = $members
        }
    
        $managedGroups.Add($add)
    }
    
    $allGroups = $null
    $myGroups = $null
    $syncHash.ManagedGroups = $managedGroups | Sort-Object Name
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

    # Listening event for Groups box
    Function Set-ListEvent {
        $count = $groupCon.SelectedItems.Count

        If ($count) {
            $item = $groupCon.SelectedItems[$count - 1]
            $details = $item.Details
        }
        Else {
            Return
        }

        $temp = [System.Collections.ArrayList]::new()
        For ($i=0;$i -lt $count;$i++) {
            $item = $groupCon.SelectedItems[$i]
            $temp.Add(($item | Select-Object -Property Members))
        }

        If ($count -gt 1) {
            If ($details) {
                $details = $details + "`r`n`r`n" + "Showing common users between selected groups below."
            }
            Else { $details = "Showing common users between selected groups below." }

            $hash = [System.Collections.Hashtable]::new()
            ForEach ($member in $temp.Members) {
                If ($hash.ContainsKey($member) -eq $false) {
                    $hash[$member] = [System.Collections.ArrayList]@($member)
                }
                Else {
                    $hash[$member].Add($member)
                }
            }

            $add = [System.Collections.ArrayList]::new()
            ForEach ($value in $hash.Values) {
                If ($value.Count -gt ($count -1)) {
                    $add.Add([pscustomobject]@{
                        Name = $value[0]
                    }) | Out-Null
                }
            }

        }
        Else {
            $add = [System.Collections.ArrayList]::new()
            ForEach ($member in $temp.Members) {
                $add.Add([pscustomobject]@{
                    Name = $member
                }) | Out-Null
            }
        }

        $detailCon.Dispatcher.Invoke(
            [action]{$detailCon.Text = $details},"Normal"
        )

        If ($syncHash.AddButton) {
            $add = ($syncHash.AddButton + $add)
        }
        
        $syncHash.RemoveButton = $null

        If ($add.Count -ne 0) { 
            $filter = ($add.Name -Replace "(?=^)|(?=$)","'$1") -Join ","
            $subtract = ($syncHash.AllUsers.Select("Name NOT IN ($filter)") | Sort-Object Name)
        }
        Else {
            $subtract = ($syncHash.AllUsers | Sort-Object Name)
        }

        If ($add.Count -gt 1) { $add = $add | Sort-Object -Property Name }
        
        $memberCon.ItemsSource = $syncHash.AddRemove = $add
        $userCon.ItemsSource = $syncHash.FilterUsers = $subtract 
    }

    # Listening event to clear everything when nothing is selected
    Function Clear-DetailsAndUsers {
        $status = $groupCon.SelectedIndex

        If ($status -eq -1) {
            $detailCon.Dispatcher.Invoke(
                [action]{$detailCon.Text = $null},"Normal"
            )

            $memberCon.ItemsSource = $syncHash.AddRemove = $null
            $userCon.ItemsSource = ($syncHash.AllUsers | Sort-Object Name)
        }
    }

    # Listening event to filter users when textbox is used
    Function Get-UserFilter {
        $data = $filterCon.Text

        If ($data) {
            $syncHash.FilterUsers = ($syncHash.AllUsers.Select("Name like '*$data*'") | Sort-Object Name)
        }
        Else {
            $syncHash.FilterUsers = ($syncHash.AllUsers | Sort-Object Name)
        }

        If ($syncHash.AddRemove -ne $null) {
            $syncHash.FilterUsers = (($syncHash.FilterUsers).Where({$_.Name -notin $syncHash.AddRemove}) | Sort-Object Name)
        }

        If ($syncHash.FilterUsers.Count -gt 1) {
            $syncHash.FilterUsers = $syncHash.FilterUsers
        }

        $userCon.ItemsSource = $syncHash.FilterUsers
    }

    # Click event for -> button
    Function Set-AddUser {
        If ($userCon.SelectedItem -eq $null) { Return }

        $data = $filterCon.Text
        $selected = ($userCon.SelectedItem | Select-Object Name)
        $addPool = [System.Collections.ArrayList]::new()

        If ($syncHash.FilterUsers -ne $null) {
            $userPool = $syncHash.FilterUsers
        }
        Else {
            $userPool = $syncHash.AllUsers
        }

        $addPool.Add($selected)

        If ($selected.Name -in $syncHash.RemoveButton.Name) {
            $syncHash.RemoveButton = ($syncHash.RemoveButton).Where({$_.Name -notin $selected.Name})
            $syncHash.AddButton = ($syncHash.AddButton).Where({$_.Name -notin $selected.Name})
        }
        Else {
            $syncHash.AddButton = $syncHash.AddButton + $addPool
        }

        If ($syncHash.AddRemove -ne $null) {
            $addPool = $addPool + $syncHash.AddRemove
        }

        If ($data) {
            $regex = ($addPool.Name | ForEach-Object {"($($_))"}) -Join "|"
            $match = ($userPool).Where({$_.Name -notmatch $regex})
            If (!$match) {
                $syncHash.FilterUsers = ($syncHash.AllUsers).Where({$_.Name -notin $syncHash.AddRemove})
            }
            Else {
                $syncHash.FilterUsers = $match
            }
        }
        Else {
            $syncHash.FilterUsers = (Compare-Object -ReferenceObject $userPool -DifferenceObject $addPool -Property Name -PassThru) | 
                Sort-Object Name
        }
        
        If ($addPool.Count -gt 1) { $addPool = $addPool | Sort-Object -Property Name }

        $userCon.ItemsSource = ($syncHash.FilterUsers | Sort-Object Name)
        $memberCon.ItemsSource = $syncHash.AddRemove = $addPool
    }

    # Click event for <- button
    Function Set-RemoveUser {
        If ($addCon.HasItems -eq $false) { Return }

        $data = $filterCon.Text
        $selected = ($memberCon.SelectedItem | Select-Object Name)
        $removePool = [System.Collections.ArrayList]::new()
        
        If ($syncHash.FilterUsers -ne $null) {
            $userPool = $syncHash.FilterUsers
        }
        Else {
            $userPool = $syncHash.AllUsers
        }

        $removePool.Add($selected)

        If ($selected.Name -in $syncHash.AddButton.Name) {
            $syncHash.AddButton = ($syncHash.AddButton).Where({$_.Name -notin $selected.Name})
            $syncHash.RemoveButton = ($syncHash.RemoveButton).Where({$_.Name -notin $selected.Name})
        }
        Else {
            $syncHash.RemoveButton = $syncHash.RemoveButton + $removePool
        }

        If ($syncHash.AddRemove.Name.Count -eq 1) { 
            $syncHash.AddRemove = $null

            $userCon.ItemsSource = ($syncHash.AllUsers | Sort-Object Name)
        }
        Else {
            $syncHash.FilterUsers = ($syncHash.FilterUsers + $removePool) | Sort-Object -Property Name
            $syncHash.AddRemove = ($syncHash.AddRemove).Where({$_.Name -notin $removePool.Name})

            $userCon.ItemsSource = ($syncHash.FilterUsers | Sort-Object Name)
        }

        If ($data) {
            $filterCon.Clear()
        }

        $memberCon.ItemsSource = $syncHash.AddRemove = ($syncHash.AddButton).Where({$_.Name -notin $removePool.Name})

    }

    # Click event for Apply button
    Function Set-Changes {
        # figure out what groups are selected and the differences between
        # group membership and the add/remove box
        $selected = $groupCon.SelectedItems

        If (($selected) -and ($syncHash.AddButton -or $syncHash.RemoveButton)) {
            # Disable the main window controls
            #$controls = @("Groups","Users","Details","AddRemove","MoveTo","MoveFrom","Apply")
            $controls = @($groupCon, $userCon, $detailCon, $memberCon, $addCon, $removeCon, $applyCon)
            ForEach ($control in $controls) {
                #$syncHash.MainWindow.FindName($control).IsEnabled = $false
                $control.IsEnabled = $false
            }

            # Prepare the 'are you sure?' message box
            $buttons = [System.Windows.MessageBoxButton]::OKCancel
            $icon    = [System.Windows.MessageBoxImage]::Information
            $title   = "Add/Remove Users"
            $body    = "Do you want to make the following changes?`r`rGroups:`r`n$($selected.Name | Out-String)`r`n"

            If ($syncHash.AddButton) {
                $body = $body + "Add Users:`r" + ($syncHash.AddButton.Name | Out-String) + "`r"
            }
            If ($syncHash.RemoveButton) {
                $body = $body + "Remove Users:`r" + ($syncHash.RemoveButton.Name | Out-String)
            }

            $result = [System.Windows.MessageBox]::Show($body,$title,$buttons,$icon)

            If ($result -eq "OK") {
                # Create a connection to the domain
                $context   = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $principal = [System.DirectoryServices.AccountManagement.PrincipalContext]::new($context)
                $idType    = [System.DirectoryServices.AccountManagement.IdentityType]::SamAccountName

                ForEach ($group in $selected) {
                    # Define the Directory Services Group
                    $sam    = $group.sAMAccountName
                    $target = [System.DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($principal,$sam)

                    # Add users to the group
                    If ($syncHash.AddButton) {
                        ForEach ($member in $syncHash.AddButton) {
                            $userSam = (($syncHash.AllUsers).Where({$_.Name -eq $member.Name})).SAM
                            $target.Members.Add($principal,$idType,$userSam)
                        }
                    }

                    # Remove users from the group
                    If ($syncHash.RemoveButton) {
                        ForEach ($member in $syncHash.RemoveButton) {
                            $userSam = (($syncHash.AllUsers).Where({$_.Name -eq $member.Name})).SAM
                            $target.Members.Remove($principal,$idType,$userSam)
                        }
                    }

                    # Save changes made to the group
                    $target.Save()
                }

                # Update the group data
                Get-ManagedGroups
                $syncHash.AddRemove = $syncHash.AddButton = $syncHash.RemoveButton = $null
                $filterCon.Clear()
                $groupCon.ItemsSource = $syncHash.ManagedGroups
                $memberCon.ItemsSource = $syncHash.AddRemove
                $userCon.ItemsSource = $syncHash.AllUsers
            }
            Else {
                # Unlock controls but do nothing
            }

            ForEach ($control in $controls) {
                #$syncHash.MainWindow.FindName($control).IsEnabled = $true
                $control.IsEnabled = $true
            }
        }
        ElseIf (($selected) -and (-Not($syncHash.AddButton -or $syncHash.RemoveButton))) {
            $buttons = [System.Windows.MessageBoxButton]::OK
            $icon    = [System.Windows.MessageBoxImage]::Information
            $title   = "Add/Remove Users"
            $body    = "There are no add or removes to perform."
            [System.Windows.MessageBox]::Show($body,$title,$buttons,$icon)
        }
        Else {
            $buttons = [System.Windows.MessageBoxButton]::OK
            $icon    = [System.Windows.MessageBoxImage]::Information
            $title   = "Add/Remove Users"
            $body    = "You have not selected any groups to perform actions on."
            [System.Windows.MessageBox]::Show($body,$title,$buttons,$icon)
        }
    }

    # Define the bare code
    [xml]$xml = '
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
                        Margin="10,5,5,5" SelectionMode="Multiple" DisplayMemberPath="Name"/>
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
                        Margin="10,5,5,5" ItemsSource="{Binding}" DisplayMemberPath="Name"/>
                </Grid> 
                <Grid Grid.Column="1">
                    <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                        <Button x:Name="MoveTo" Content="&#x1f846;" Style="{StaticResource ButtonStyle}"/>
                        <Button x:Name="MoveFrom" Content="&#x1f844;" Style="{StaticResource ButtonStyle}"/>
                    </StackPanel>
                </Grid>
                <Grid Grid.Column="2">
                <ListBox x:Name="AddRemove" HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                    Margin="5,5,10,5" DisplayMemberPath="Name"/>
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

    # Define the main window
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $syncHash.MainWindow = [Windows.Markup.XamlReader]::Load($reader)

    # Define the controls for greater ease
    $script:groupCon  = $syncHash.MainWindow.FindName("Groups")
    $script:logCon    = $syncHash.MainWindow.FindName("LoggedIn")
    $script:detailCon = $syncHash.MainWindow.FindName("Details")
    $script:userCon   = $syncHash.MainWindow.FindName("Users")
    $script:addCon    = $syncHash.MainWindow.FindName("MoveTo")
    $script:removeCon = $syncHash.MainWindow.FindName("MoveFrom")
    $script:memberCon = $syncHash.MainWindow.FindName("AddRemove")
    $script:filterCon = $syncHash.MainWindow.FindName("Filter")
    $script:applyCon  = $syncHash.MainWindow.FindName("Apply")

    # ...and in this area, bind them.
    $groupCon.ItemsSource = $syncHash.ManagedGroups
    $userCon.ItemsSource = ($syncHash.AllUsers | Sort-Object Name)
    $logCon.Text = "Logged in as: $($env:USERNAME)"

    # Add listeners
    $groupCon.Add_SelectionChanged({ Set-ListEvent; Clear-DetailsAndUsers })
    $filterCon.Add_TextChanged({ Get-UserFilter })
    $addCon.Add_Click({ Set-AddUser })
    $removeCon.Add_Click({ Set-RemoveUser })
    $applyCon.Add_Click({ Set-Changes })

    Close-Window
    [void]$syncHash.MainWindow.ShowDialog()
}

# ----------------------------------------------------------------------------------------------------------- #
# Start application; show user a progress window
Show-ProgressWindow
Set-ProgressMessage -message "Gathering group information"

# Get bulk information from AD
$syncHash.AllUsers = Get-AllUsers
Get-ManagedGroups

# Close progress window and open main form
Show-MainWindow