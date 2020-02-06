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

Function Start-ProcessingListItems {
    Param(
        $collection,
        $name,
        $element
    )

    # Based on the amount of total objects, figure out how to divide them
    # into smaller groups for processing
    $count = ($syncHash.$collection).Count

    If ($count -le 100) { $division = 10 }
    If (($count -gt 101) -and ($count -le 500)) { $division = 25 }
    If (($count -gt 501) -and ($count -le 1000)) { $division = 50 }
    If ($count -gt 1001) { $division = 100 }
    
    # create array of work cycles
    $groups = [System.Collections.ArrayList]::new()

    # If the count doesn't divide evenly, account for the remainder
    If ($count % $division) {
        $a = [math]::truncate($count / $division)
        $b = ($count % $division)
    } 
    Else {
        $a = ($count / $division)
    }

    For ($i=1;$i -lt ($a + 1);$i++){
        $total = $i * $division
        $start = ($total - $division)
        $end = ($total - 1)
        $temp = [pscustomobject]@{
            Group = $i
            Start = $start
            End = $end
        }
        $groups.Add($temp) | Out-Null
    }
    
    If ($b) {
        $temp = [pscustomobject]@{
            Group = $i
            Start = ($end + 1)
            End = ($end + $b)
        }
        $groups.Add($temp) | Out-Null
    }

    # Create array for monitoring all the runspaces
    $runspaceCollection = @()

    # Define some functions that need to be included in the runspaces
    $n = 0
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $variableEntryA = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'syncHash',$syncHash,$null

    # Add the synchronized table to the runspaces
    $initialSessionState.Variables.Add($variableEntryA)

    # Create a runspace pool to process all of the information
    $runspacePool = [RunspaceFactory]::CreateRunspacePool(1,20,$initialSessionState,$host)
    $runspacePool.ApartmentState = "STA"
    $runspacePool.ThreadOptions  = "ReuseThread"
    $runspacePool.Open()

    # Define what the runspaces will be doing
    $scriptBlock = {
        Param($start, $end, $name, $element, $collection)

        For ($i=$start;$i -lt ($end+1);$i++) {
            $item = ($syncHash.$collection).Name[$i]

            $newListItem = [pscustomobject]@{
                Name    = $name+$i 
                Content = $item
            }
            #>
            #$newListItem = "<ListBoxItem x:Name=`"$name$i`" Content=`"$item`" />`r`n"
            $syncHash."Add$element".Add($newListItem)
        }
        
    }

    # Begin processing everything
    # begin working
    While (!$complete) {

        $start = $groups[$n].Start
        $end = $groups[$n].End
        $parameters = @{
            start = $start
            end = $end
            name = $name
            element = $element
            collection = $collection
        }

        # There can only be 10 jobs running at any given time
        # If there's less than 10, add a job
        If (($runspaceCollection.Count -le 20) -and ($n -lt $groups.Count)) {
            # Create the powershell object that's going to run the job
            $powershell = [powershell]::Create().AddScript($scriptblock).AddParameters($parameters)

            # Add the powerhshell job to the pool
            $powershell.RunspacePool = $runspacePool

            # Add monitoring to the runspace collection and start the job
            [collections.arraylist]$runspaceCollection += new-object psobject -property @{
                Runspace = $powershell.BeginInvoke()
                PowerShell = $powershell
            }

            # Iterate n
            $n++
        }

        # Check the job status and post results
        ForEach ($runspace in $runspaceCollection.ToArray()) {
            If ($runspace.Runspace.IsCompleted) {
                # Remove the runspace so a new one can be built
                $runspace.PowerShell.Dispose()
                $runspaceCollection.Remove($runspace)
            }
        }

        # Define the complete parameter
        if (($n -eq $groups.Count) -and ($runspaceCollection.Count -eq 0)){
            $complete = $true
        }

    }

    # Close and dispose of the pool
    $runspacePool.Close()
    $runspacePool.Dispose()
    
}

Function Show-MainWindow {

    Function GroupList-Event {
        $selected = $this
        $i = $selected.Name -Replace "Group"
        $details = ($syncHash.ManagedGroups[$i]).Details
        $members = ($syncHash.ManagedGroups[$i]).Members
        $selected = $syncHash.MainWindow.FindName("Groups").SelectedItems.Count

        $syncHash.MainWindow.FindName("Details").Dispatcher.Invoke(
            [action]{$syncHash.MainWindow.FindName("Details").Text = $details},"Normal"
        )

        If ($selected -gt 0) {
            
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
        #>
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
                        Margin="10,5,5,5" SelectionMode="Multiple">
                        <!-- Replace Groups -->
                    </ListBox>
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
                        Margin="10,5,5,5">
                        <!-- Replace Users -->
                    </ListBox>
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

    # Add Groups to the Group element
    Set-ProgressMessage -message "Loading groups to window"
    Start-ProcessingListItems -Collection "ManagedGroups" -Name "Group" -Element "Groups"
    
    $temp = ($syncHash.AddGroups | Sort-Object -Property Content)
    ForEach ($item in $temp) {
        $child = $xml.CreateElement("ListBoxItem", $xml.Window.NamespaceURI)
        $child.SetAttribute("x:Name",$item.Name)
        $child.SetAttribute("Content",$item.Content)
        $xml.GetElementsByTagName("ListBox")[0].AppendChild($child) | Out-Null
    }

    # Add Users to the User element
    Set-ProgressMessage -message "Loading users to window"
    Start-ProcessingListItems -Collection "AllUsers" -Name "User" -Element "Users"

    $temp = ($syncHash.AddUsers | Sort-Object -Property Content)
    ForEach ($item in $temp) {
        If ($item.Name) {
            $child = $xml.CreateElement("ListBoxItem", $xml.Window.NamespaceURI)
            $child.SetAttribute("x:Name",$item.Name)
            $child.SetAttribute("Content",$item.Content)
            $xml.GetElementsByTagName("ListBox")[1].AppendChild($child) | Out-Null
        }
    }

    # define the main window
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $syncHash.MainWindow = [Windows.Markup.XamlReader]::Load($reader)

    # Add a listener for each ListBoxItem
    $listBoxItems = $syncHash.MainWindow.FindName("Groups").Items
    ForEach ($listBoxItem in $listBoxItems) {
        $listBoxItem.Add_Selected({ GroupList-Event })
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
$syncHash.AllUsers = Get-AllUsers | Sort-Object Name
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
$syncHash.ManagedGroups = $managedGroups | Sort-Object Name
$syncHash.AddGroups = [System.Collections.ArrayList]::new()
$syncHash.AddUsers = [System.Collections.ArrayList]::new()

# Close progress window and open main form
Show-MainWindow