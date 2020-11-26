$archives = "D:\SIEM Data\2019\Aug\Stibo_Archives"#Change this to Cab path#

Function Invoke-MDBSQLCMD ($mdblocation,$sqlquery){
$dsn = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=$mdblocation;"
$objConn = New-Object System.Data.OleDb.OleDbConnection $dsn
$objCmd  = New-Object System.Data.OleDb.OleDbCommand $sqlquery,$objConn
$objConn.Open()
$adapter = New-Object System.Data.OleDb.OleDbDataAdapter $objCmd
$dataset = New-Object System.Data.DataSet
[void] $adapter.Fill($dataSet)
$objConn.Close()
$dataSet.Tables | Select-Object -Expand Rows
$dataSet = $null
$adapter = $null
$objCmd  = $null
$objConn = $null
}

Install-Module -Name '7Zip4Powershell'

(Get-ChildItem -path $archives\*.cab)| ForEach-Object {
Expand-7Zip -ArchiveFileName $_.FullName -TargetPath $archives
}


$regex2 = '\%ASA\-\d\-(\d+)\:'
Filter Extract2 {
$_.event_description -match $regex2 > $null
[pscustomobject]@{  
MsgID = ($Matches[1])
}}

(Get-ChildItem -path $archives\*.mdb)| ForEach-Object {
$mdblocation1 = $_.fullname
$mdbname1 = $_.basename
$fname= "CountDevices_{0}" -f $mdbname1
$query1 = (Invoke-MDBSQLCMD $mdblocation1 -sqlquery "Select computer from Events_Table")
$result = $query1 | Group-Object computer | Sort-Object Count -Descending | Select-Object -Property @{name="DeviceName";expression={$($_.Name)}}, Count
$result | Export-Csv -Path "$archives\$fname.csv" -NoTypeInformation

$fname1= "CountMessageID_{0}" -f $mdbname1
$query2 = ((Invoke-MDBSQLCMD $mdblocation1 -sqlquery "Select event_description from Events_Table")| Extract2)
$result2 = $query2 | Group-Object MsgID | Sort-Object Count -Descending | Select-Object -Property @{name="MessageID";expression={$($_.Name)}}, Count
$result2 | Export-Csv -Path "$archives\$fname1.csv" -NoTypeInformation
}

Get-ChildItem -Path $archives\CountMessageID_*.csv | # Get each CSV files
    ForEach-Object -Process {
        Import-Csv -Path $PSItem.FullName # Import CSV data
    } | 
    Group-Object -Property MessageID | # Group per Domain Name
    Select-Object -Unique -Property Name, @{
        Label = "Sum";
        Expression = {
            # Sum all the counts for each domain
            ($PSItem.group | Measure-Object -Property Count -sum).Sum
        }
    } |
    Sort-Object -Property Sum -Descending| Select-Object -Property @{name="CumulativeMessageID";expression={$($_.Name)}},@{name="CumulativeCount";expression={$($_.Sum)}}  | Export-Csv -Path "$archives\Cumulative_MsgID_Count.csv" -NoTypeInformation

Get-ChildItem -Path $archives\CountDevices_*.csv | # Get each CSV files
    ForEach-Object -Process {
        Import-Csv -Path $PSItem.FullName # Import CSV data
    } | 
    Group-Object -Property DeviceName | # Group per Domain Name
    Select-Object -Unique -Property Name, @{
        Label = "Sum";
        Expression = {
            # Sum all the counts for each domain
            ($PSItem.group | Measure-Object -Property Count -sum).Sum
        }
    } |
    Sort-Object -Property Sum -Descending| Select-Object -Property @{name="CumulativeDeviceName";expression={$($_.Name)}},@{name="CumulativeCount";expression={$($_.Sum)}}  | Export-Csv -Path "$archives\Cumulative_Device_Count.csv" -NoTypeInformation

Remove-Item -Path $archives\*.csv -Exclude "Cumulative_*.csv"
Remove-Item -Path $archives\*.mdb
