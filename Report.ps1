#Import-Module VMware.PowerCLI
#Set-PowerCLIConfiguration -InvalidCertificateAction Ignore
#connect-omserver 192.168.10.151

$excel = New-Object -ComObject excel.application 
$excel.visible = $True
$workbook = $excel.Workbooks.Add()
$workbook.Worksheets.Item(2).Delete()
$computewrksht= $workbook.Worksheets.Item(1) 
$computewrksht.Name = 'Compute'
$networkwrksht= $workbook.Worksheets.Item(2) 
$networkwrksht.Name = 'Network'
$computewrksht.Cells.Item(1,1)= 'NonProd'
$computewrksht.Cells.Item(1,1).ColumnWidth = 13.86
$computewrksht.Cells.Item(1,2).ColumnWidth =14.48
$computewrksht.Cells.Item(1,3).ColumnWidth =8.43
$computewrksht.Cells.Item(1,4).ColumnWidth =8.43
$computewrksht.Cells.Item(1,5).ColumnWidth =8.43
$computewrksht.Cells.Item(1,6).ColumnWidth =8.43
$computewrksht.Cells.Item(1,7).ColumnWidth =19.41

$MergeCells = $computewrksht.Range("A1:G1") 
$MergeCells.Select() 
$MergeCells.MergeCells = $true 
$computewrksht.Cells.Item(1, 1).HorizontalAlignment = -4108
$computewrksht.Cells.Item(1,1).Font.Size = 11
$computewrksht.Cells.Item(1,1).Font.Bold=$True 
$computewrksht.Cells.Item(1,1).Font.Name = "Calibri" 
$computewrksht.Cells.Item(1,1).Font.ColorIndex = 2
$computewrksht.Cells.Item(1,1).interior.colorindex=1



$clusters = Get-OMResource -resourcekind ClusterComputeResource

$clusterName = "test"
$row=1

foreach($clusterName in $clusters){

$row++
$MergeCells = $computewrksht.Range($computewrksht.Cells.Item($row,1),$computewrksht.Cells.Item($row,7)) 
$MergeCells.Select() 
$MergeCells.MergeCells = $true 
$computewrksht.Cells.Item($row,1)= $clusterName.Name
$computewrksht.Cells.Item($row, 1).HorizontalAlignment = -4108
$computewrksht.Cells.Item($row,1).Font.Size = 11
$computewrksht.Cells.Item($row,1).Font.Bold=$True 
$computewrksht.Cells.Item($row,1).Font.Name = "Calibri" 
$computewrksht.Cells.Item($row,1).interior.colorindex=16
$row++
$range = $computewrksht.Range($computewrksht.Cells.Item($row,1),$computewrksht.Cells.Item($row+3,7))
$range.HorizontalAlignment = -4152
$computewrksht.Cells.Item($row,1).RowHeight = 30
$computewrksht.Cells.Item($row,2) = "Total capacity"
$computewrksht.Cells.Item($row,3) = "Demand"
$computewrksht.Cells.Item($row,4) = "Peak"
$computewrksht.Cells.Item($row,5) = "% Avg Demand"
$computewrksht.Cells.Item($row,6) = "% Peak Demand"
$computewrksht.Cells.Item($row,7) = "Time Remaining"
$computewrksht.Cells.Item($row,5).wraptext = $true
$computewrksht.Cells.Item($row,6).wraptext = $true

$row++
$computewrksht.Cells.Item($row,1) = "CPU"
$computewrksht.Cells.Item($row+1,1) = "Memory"
$computewrksht.Cells.Item($row+2,1) = "Storage"
$computewrksht.Cells.Item($row,2) =  ($clusterName | Get-OMStat -Key "cpu|capacity_provisioned" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,3) =  ($clusterName | Get-OMStat -Key "cpu|demandmhz" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,4) =  ($clusterName | Get-OMStat -Key "cpu|demandmhz" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,5) =  ($clusterName | Get-OMStat -Key "cpu|demandPct" -RollupType "Avg" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,6) =  ($clusterName | Get-OMStat -Key "cpu|demandPct" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,7) =  ($clusterName | Get-OMStat -Key "OnlineCapacityAnalytics|cpu|demand|timeRemaining" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$row++
$computewrksht.Cells.Item($row,2) =  ($clusterName | Get-OMStat -Key "mem|host_provisioned" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,3) =  ($clusterName | Get-OMStat -Key "mem|host_usage" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,4) =  ($clusterName | Get-OMStat -Key "mem|host_usage" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,5) =  ($clusterName | Get-OMStat -Key "mem|host_usagePct" -RollupType "Avg" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,6) =  ($clusterName | Get-OMStat -Key "mem|host_usagePct" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,7) =  ($clusterName | Get-OMStat -Key "OnlineCapacityAnalytics|mem|demand|timeRemaining" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value

$row++

$computewrksht.Cells.Item($row,2) =  ($clusterName | Get-OMStat -Key "diskspace|total_capacity" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,3) =  ($clusterName | Get-OMStat -Key "diskspace|total_usage" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,4) =  ($clusterName | Get-OMStat -Key "diskspace|total_usage" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,5) =  ($clusterName | Get-OMStat -Key "diskspace|demand|workload" -RollupType "Avg" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,6) =  ($clusterName | Get-OMStat -Key "diskspace|demand|workload" -RollupType "Max" -IntervalType "Weeks" -IntervalCount 1).Value
$computewrksht.Cells.Item($row,7) =  ($clusterName | Get-OMStat -Key "OnlineCapacityAnalytics|diskspace|demand|timeRemaining" -RollupType "Latest" -IntervalType "Weeks" -IntervalCount 1).Value

}


#$workbook.SaveAs($outputpath) 
#$excel.Quit()