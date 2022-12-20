################################################################################################################################################
#
#    Program Name:     POTA_Online.ps1
#    Description:      This PowerShell script provides SMS Text and email notifications when certain criteria from a "WatchList" has
#                      been met.  The Watchlist file is a .csv file that can be managed by either Microsoft Excel or Notepad.
#    Date:             12/15/2022
#    Author:           Kurt W. Zimmerman - W2MW
#    Revisions:        12/15/2022 - Initial Release
#                      12/20/2022 - Added "Ignored" to Watchlist
#                                   Added Quiet Time logic to not send out SMS messages during certain hours of the day
#                                   Added N3FJP lookup to see if station that met WatchList criteria has been worked.  If so then ignore
#                                    notification.
#                                   Added update to N3FJP to populate Park and Park description to QSO in log.
#
################################################################################################################################################

####
#     Output File information
####
$Debug = $False
$Enable_Quite_Time = $true

####
#     Set hours of no notifications
#     Then check to see if program should run based on current time
####
$QuiteTimes_HH = (23, 24, 1, 2, 3, 4, 5, 6) -split ","
[int]$Current_Time = get-date -format "HH"

$IS_Quiet = $false
foreach ($q_HH in $QuiteTimes_HH) {
    if ($Current_Time -eq $q_HH) {
        $IS_Quiet = $true
        break
    }
}

if (($Enable_Quite_Time -and !$IS_Quiet) -or (!$Enable_Quite_Time)) {
    ####
    #
    #  Program Setup 
    #
    ####
    $Export_Raw_File = 'D:\Transfer\POTA\Pota_Raw.csv'
    $Export_Fiter_File = 'D:\Transfer\POTA\Pota_Filter.csv'
    $Export_Log_File = 'D:\Transfer\POTA\Pota_Log.csv'
    $WatchList_File = 'D:\Transfer\POTA\WatchList.csv'
    $Export_Raw_File = 'D:\Transfer\POTA\Pota_Raw.txt'
    $N3FJP_Database = 'C:\Users\Our Place\Dropbox\PC\Documents\Affirmatech\N3FJP Software\ACLog\LogData.mdb'
    $WatchList = Import-Csv -Path $WatchList_File | Where-Object { $_.Ignore -ne 'Yes' }
    $CRLF = "`r`n"
    $Launch_Excel = $False
    $update_N3FJP = $true
    $UTCDate = get-date -format "yyyy/MM/dd" (get-date).ToUniversalTime()
    $TodayDateTime = (get-date -format "MM/dd/yyyy HH:mm").ToString()
    #     $TodayDate = (get-date -format "yyyy/MM/dd")
    ####
    #    POTA stations who are on
    ####
    $WebResponse = Invoke-WebRequest "https://api.pota.app/spot/activator"
    $WebResponse = ((($WebResponse -replace ('{', '') -replace ('}', '')) -replace ('\[', '')) -replace ('\]', ''))
    $POTA_Spots = $WebResponse -split ','
    $PotaTable = New-Object System.Data.DataTable
    $PotaTable_Output = New-Object System.Data.DataTable
    $POTA_Spots | out-file -filepath $Export_Raw_File

    ####
    #     Email information
    ####
    $From = "<email address>@gmail.com"
    $PW = "<password>"
    $To = "<To Email Address"
    $SMS_Address = '<SMS Phone Number>@mms.att.net'
    $global:SMS_Body = ''
    $Subject = "POTA stations online"
    $Body = "These stations are presently online"
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    $SentSMS = $False

    $passwordAD = ConvertTo-SecureString $PW -AsPlainText -Force
    $CredAD = New-Object System.Management.Automation.PSCredential ($From, $passwordAD)

    ####
    #     Building Band Plan Table
    ####
    $BandPlan = New-Object System.Data.DataTable
    $BandPlan.Columns.Add("Band", "System.Decimal") | Out-Null
    $BandPlan.Columns.Add("Lower", "System.Decimal") | Out-Null
    $BandPlan.Columns.Add("Upper", "System.Decimal") | Out-Null

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 160
    $BandPlanRow.Lower = 1.8
    $BandPlanRow.Upper = 2.0
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 80
    $BandPlanRow.Lower = 3.5
    $BandPlanRow.Upper = 4.0
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 60
    $BandPlanRow.Lower = 5.3305
    $BandPlanRow.Upper = 5.4035
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 40
    $BandPlanRow.Lower = 7.0
    $BandPlanRow.Upper = 7.3
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 30
    $BandPlanRow.Lower = 10.1
    $BandPlanRow.Upper = 10.15
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 20
    $BandPlanRow.Lower = 14.0
    $BandPlanRow.Upper = 14.35
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 17
    $BandPlanRow.Lower = 18.068
    $BandPlanRow.Upper = 18.168
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 15
    $BandPlanRow.Lower = 21.0
    $BandPlanRow.Upper = 21.45
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 12
    $BandPlanRow.Lower = 24.89
    $BandPlanRow.Upper = 24.99
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 10
    $BandPlanRow.Lower = 28
    $BandPlanRow.Upper = 29
    $BandPlan.Rows.Add($BandPlanRow)

    $BandPlanRow = $BandPlan.NewRow()
    $BandPlanRow.Band = 6
    $BandPlanRow.Lower = 50
    $BandPlanRow.Upper = 54
    $BandPlan.Rows.Add($BandPlanRow)

    ####
    # Lookup Band Function
    ####
    Function Lookup_Band ([decimal]$InFrequency) {
        [string]$CurrentBand = 'Unknown'
        foreach ($band in $BandPlan) {
            if (($Infrequency -ge $band.lower) -and ($Infrequency -le $band.upper)) {
                $CurrentBand = $band.Band.tostring()
                break
            }
        }
        return $CurrentBand
    }
    ####
    # Lookup QSO in N3FJP log Function
    ####
    Function Lookup_QSO {
        Param(
            [string]$fileName  
            , [string]$fldCall
            , [string]$fldBand
            , [string]$fldMode
            , [string]$fldDateStr
            , [string]$fldOther2
            , [string]$fldOther3
        )
    
        $conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filename;Persist Security Info=False")
        $cmd = $conn.CreateCommand()
        $cmd.CommandText
        $SelectCmd = "SELECT [fldPrimaryKey]
        ,[fldDateStr]
        ,[fldBand]
        ,[fldMode]
        ,[fldCall]
        ,[fldFrequency]
        ,[fldOther2]
        ,[fldOther3]
        ,[fldComments]
    FROM [tblContacts]
    where fldCall = '" + $fldCall + "'" + $CRLF
        if ($fldBand) { $SelectCmd += ' and fldBand = "' + $fldBand + '"' + $CRLF }
        if ($fldMode) { $SelectCmd += ' and fldMode = "' + $fldMode + '"' + $CRLF }
        if ($fldDateStr) { $SelectCmd += ' and fldDateStr = "' + $fldDateStr + '"' + $CRLF }
        #$SelectCmd += " and fldOther2 = switch ( fldOther2 = '', fldOther2,
        #                                         fldOther2 <> '', '" + $fldOther2 + "')" + $CRLF
        $SelectCmd += ' order by fldPrimaryKey desc'
        $cmd.CommandText = $SelectCmd
        $conn.open()
        $rdr = $cmd.ExecuteReader()
        $global:QSOs = New-Object System.Data.Datatable
        $global:QSOs.Load($rdr)
        $conn.close()
    }

    Function Update_QSO {
        Param(
            [string]$fileName  
            , [string]$fldPrimaryKey
            , [string]$fldOther2
            , [string]$fldOther3
        )
    
        $conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filename;Persist Security Info=False")
        $cmd = $conn.CreateCommand()
        $cmd.CommandText
        $SelectCmd = "update tblContacts
                 set fldOther2 = '" + $fldOther2 + "', fldOther3 = '" + $fldOther3 + "'  where fldPrimaryKey = " + $fldPrimaryKey.ToString()
        $cmd.CommandText = $SelectCmd
        $conn.open()
        #$rdr = $cmd.ExecuteReader()
        $cmd.ExecuteReader()
        $conn.close()
    }

    #
    #  Function to populate output tables and SMS message
    #

    function AddOutputRow ($Matched) {
        $PotaTableRow = $PotaTable_Output.NewRow()
        $PotaTableRow.spotId	= $PotaTable_Row.spotId
        $PotaTableRow.activator = $PotaTable_Row.activator
        $PotaTableRow.frequency	= $PotaTable_Row.frequency
        $PotaTableRow.mode	= $PotaTable_Row.mode
        $PotaTableRow.reference	= $PotaTable_Row.reference
        $PotaTableRow.parkName	= $PotaTable_Row.parkName
        $PotaTableRow.spotTime	= $PotaTable_Row.spotTime
        $PotaTableRow.spotter	= $PotaTable_Row.spotter
        $PotaTableRow.comments	= $PotaTable_Row.comments
        $PotaTableRow.source	= $PotaTable_Row.source
        $PotaTableRow.invalid	= $PotaTable_Row.invalid
        $PotaTableRow.name = $PotaTable_Row.name
        $PotaTableRow.locationDesc	= $PotaTable_Row.locationDesc
        $PotaTableRow.grid4	= $PotaTable_Row.grid4
        $PotaTableRow.grid6	= $PotaTable_Row.grid6
        $PotaTableRow.latitude	= $PotaTable_Row.latitude
        $PotaTableRow.longitude	= $PotaTable_Row.longitude
        $PotaTableRow.count	= $PotaTable_Row.count
        $PotaTableRow.expire = $PotaTable_Row.expire
        $PotaTableRow.matched = $Matched
        $PotaTable_Output.Rows.Add($PotaTableRow)

        $global:SMS_Body = 'Activator: ' + $PotaTable_Row.activator + $CRLF
        $global:SMS_Body += 'Frequency : ' + $PotaTable_Row.frequency + $CRLF
        $global:SMS_Body += 'Mode : ' + $PotaTable_Row.mode + $CRLF
        $global:SMS_Body += 'Reference: ' + $PotaTable_Row.reference + $CRLF 
        $global:SMS_Body += 'Park Nane: ' + $PotaTable_Row.name + $CRLF
        $global:SMS_Body += 'Spot Time: ' + (get-date $PotaTable_Row.spotTime -format "MM/dd/yyyy HH:mm").ToString() + $CRLF
        $global:SMS_Body += 'Current Date/Time ' + $TodayDateTime + $CRLF 
        $global:SMS_Body += 'locationDesc: ' + $PotaTable_Row.locationDesc + $CRLF
        $global:SMS_Body += 'Comments: ' + $PotaTable_Row.comments + $CRLF
        $global:SMS_Body += $Matched + $CRLF
        $global:SMS_Body += '- - - -' + $CRLF       
    }

    ####   
    # 
    #  Start of main program
    #
    #  Building source table
    #
    ####
    foreach ($POTA_Spots_Detail in $POTA_Spots) {
        $PotaColumnMatch = $False
        $Pota_Spot_Data = $POTA_Spots_Detail.trim() -replace ('"', '')
        if ($Pota_Spot_Data -like '*:*') {
            $index = $Pota_Spot_Data.IndexOf(":")
            $Pota_Spot_Data_Column = $Pota_Spot_Data.trim() -replace (':', '')
            $Pota_Spot_Data_Column = $Pota_Spot_Data_Column.substring( 0, $Index)
            if ($Debug) {
                $Pota_Spot_Data_Column
            }
            foreach ($PotaColumn in $PotaTable.Columns.ColumnName) {
                IF ($PotaColumn -eq $Pota_Spot_Data_Column) {
                    $PotaColumnMatch = $true
                    break
                }
            }
            if (!$PotaColumnMatch) {
                $PotaTable.Columns.Add($Pota_Spot_Data_Column) | Out-Null
                $PotaTable_Output.Columns.Add($Pota_Spot_Data_Column) | Out-Null
            }
        }
    }
    $PotaTable.Columns.Add("Matched") | Out-Null
    $PotaTable_Output.Columns.Add("Matched") | Out-Null

    #
    # Loading data into table
    #
    $RowAdded = $False
    foreach ($POTA_Spots_Detail in $POTA_Spots) {
        $Pota_Spot_Data = $POTA_Spots_Detail.trim() -replace ('"', '')
        if ($Pota_Spot_Data -like '*:*') {
            $PostData = $false
            $index = $Pota_Spot_Data.IndexOf(":")
            $Pota_Spot_Data_Column = $Pota_Spot_Data.trim() -replace (':', '')
            $Pota_Spot_Data_Column = $Pota_Spot_Data_Column.substring( 0, $Index)
            $Pota_Spot_Data_Value = $POTA_Spots_Detail.substring($Index + 4 )
            if ($Pota_Spot_Data -like 'spotId:*') {
                if ($RowAdded) {
                    $PotaTable.Rows.Add($PotaRow) 
                    $RowAdded = $False
                }
                $PotaRow = $PotaTable.NewRow()
                $RowAdded = $true
            }
            foreach ($PotaColumn in $PotaTable.Columns.ColumnName) {
                IF ($PotaColumn -eq $Pota_Spot_Data_Column) {
                    if ($Debug) {
                        $PotaColumn
                        $Pota_Spot_Data_Value
                    }
                    $PotaRow[$PotaColumn] = ($Pota_Spot_Data_Value.Replace('"', '')).trim()
                    $PostData = $true
                    #break
                }
            }
            if (!$PostData) {
                write-host $Pota_Spot_Data_Column
            }
        }
    }
    if ($RowAdded) {
        $PotaTable.Rows.Add($PotaRow) 
    }
    ####
    #
    #  Loop through Pota Spots
    #
    ####
    [string]$Matched = ''
    foreach ($PotaTable_Row in $PotaTable) {
        $CurrentBand = Lookup_Band -InFrequency ($PotaTable_Row.frequency / 1000)
        ####
        #
        #  Now loop through WatchList to see if there is a match
        #
        ####
        foreach ($WatchList_Row in $WatchList) {
            if ($Debug) {
                $WatchList_Row
            }
            $FoundSpot1 = $false
            $FoundSpot2 = $false
            $FoundSpot3 = $false
            $FoundSpot4 = $false
            $FoundSpot5 = $false
            $FoundSpot6 = $False
            $FoundQRT = $False
            $Matched = ''
            if ($Debug) {
                Write-host ($PotaTable_Row.activator.ToString() + ' | ' + $WatchList_Row.activator.ToString())
            }
            if (($WatchList_Row.activator -eq $PotaTable_Row.activator) -or $WatchList_Row.activator -eq '*') {
                if ($WatchList_Row.activator -ne '*') {
                    $Matched = 'Match on: Activator - ' + $PotaTable_Row.activator 
                }
                $FoundSpot1 = $true
            }
            if ($Debug) { 
                Write-host ($PotaTable_Row.frequency.tostring() + ' | ' + $WatchList_Row.frequency.ToString()) 
            }
        
            if ($WatchList_Row.frequency -eq $PotaTable_Row.frequency -or $WatchList_Row.frequency -eq '*') {
                if ($WatchList_Row.frequency -ne '*') {
                    if ($Matched) { $Matched += ' ' + $CRLF }
                    $Matched += 'Match on: Frequency - ' + $PotaTable_Row.frequency 
                }
                $FoundSpot2 = $true
            }
            if ($debug) {
                Write-host ($PotaTable_Row.mode.tostring() + ' | ' + $WatchList_Row.mode.ToString()) 
            }

            if ($WatchList_Row.mode -eq $PotaTable_Row.mode -or $WatchList_Row.mode -eq '*') {
                if ($WatchList_Row.mode -ne '*') {
                    if ($Matched) { $Matched += ' ' + $CRLF }
                    $Matched += 'Match on: Mode - ' + $PotaTable_Row.mode
                }
                $FoundSpot3 = $true
            }
            if ($Debug) {
                Write-host ($PotaTable_Row.reference.tostring() + ' | ' + $WatchList_Row.reference.ToString()) 
            }

            if ($WatchList_Row.reference -eq $PotaTable_Row.reference -or $WatchList_Row.reference -eq '*') {
                if ($WatchList_Row.reference -ne '*') {
                    if ($Matched) { $Matched += ' ' + $CRLF }
                    $Matched += 'Match on: Reference - ' + $PotaTable_Row.reference
                }
                $FoundSpot4 = $true
            }
            if ($Debug) {
                Write-host ($PotaTable_Row.locationDesc.tostring() + ' | ' + $WatchList_Row.locationDesc.ToString()) 
            }

            if ($WatchList_Row.locationDesc -eq $PotaTable_Row.locationDesc -or $WatchList_Row.locationDesc -eq '*') {
                if ($WatchList_Row.locationDesc -ne '*') {
                    if ($Matched) { $Matched += ' ' + $CRLF }
                    $Matched += 'Match on: locationDesc (Country-State) - ' + $PotaTable_Row.locationDesc
                }                
                $FoundSpot5 = $true
            }

            if ($Debug) {
                Write-host ($CurrentBand.tostring() + ' | ' + $WatchList_Row.Band.ToString()) 
            }

            if ($WatchList_Row.Band.ToString() -eq $CurrentBand -or $WatchList_Row.Band -eq '*') {
                if ($WatchList_Row.Band -ne '*') {
                    if ($Matched) { $Matched += ' ' + $CRLF }
                    $Matched += 'Match on: Band - ' + $CurrentBand.ToString()
                }                
                $FoundSpot6 = $true
            }
            if ($Debug) {
                write-host $PotaTable_Row.comments 
            }

            if ($PotaTable_Row.comments -like '*qrt*') {
                $FoundQRT = $true
            }
      
            if ($FoundSpot1 -and $FoundSpot2 -and $FoundSpot3 -and $FoundSpot4 -and $FoundSpot5 -and $FoundSpot6 -and !$FoundQRT) {
                ####
                #
                #  Now that there is a match, look in the log to see if POTA Station has already been worked.
                #
                ####
                Lookup_QSO -fileName $N3FJP_Database -fldCall $PotaTable_Row.activator -fldBand $CurrentBand -fldMode $PotaTable_Row.mode -fldDateStr $UTCDate -fldOther2 $PotaTable_row.reference -fldOther3 $PotaTable_row.Name
                if (!$global:QSOs -or (($Global:QSOs.fldOther2 -ne $PotaTable_Row.reference) -and $Global:QSOs.fldOther2 -ne '')) {
                    #
                    #  Build output table and SMS message
                    #
                    AddOutputRow ($Matched)
        
                    $PotaTable_Output | Export-Csv -Path $Export_Fiter_File -NoTypeInformation
                    $PotaTable_Output | Export-Csv -Path $Export_Log_File -Append -NoTypeInformation
                    $PotaTable | Export-Csv -Path $Export_Raw_File -NoTypeInformation
        
                    #
                    # Sending out SMS Alert
                    #
                    Send-MailMessage -From $From -to $SMS_Address -Subject $Subject `
                        -Body $global:SMS_Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
                        -Credential $CredAD 
                    $SentSMS = $True
                }
                #
                # Update Park Reference and Park Name if not updated
                #
                elseif ((($global:QSOs -and $Global:QSO.fldOther2 -ne $PotaTable_Row.reference)) -and $update_N3FJP) {
                    Update_QSO -fileName $N3FJP_Database -fldPrimaryKey $global:QSOs.fldPrimaryKey -fldOther2 $PotaTable_Row.reference -fldOther3 $PotaTable_Row.Name
                }
            }
        }
    }
    If ($SentSMS) {
        ####
        #
        # Only send out Email if SMS alert was sent 
        #
        ####
        Send-MailMessage -From $From -to $To -Subject $Subject `
            -Body $Body -Attachments $Export_Fiter_File, $Export_Raw_File, $WatchList_File -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
            -Credential $CredAD 
    }

    if ($Launch_Excel) {
        # start Excel
        $excel = New-Object -comobject Excel.Application
        #open file
        #$FilePath = 'D:\Transfer\POTA\Book1.xlsm'
        #$workbook = $excel.Workbooks.Open($Export_Fiter_File)
        #$workbook = $excel.Workbooks.Open($Export_Raw_File)

        $excel.Workbooks.Open($Export_Fiter_File)
        $excel.Workbooks.Open($Export_Raw_File)
        #make it visible (just to check what is happening)
        $excel.Visible = $true
        #access the Application object and run a macro
        #$app = $excel.Application 
        $excel.Application 
        #$app.Run("Macro1")
    }
}