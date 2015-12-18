
Sign in to vote
$objExcel = New-Object -comobject Excel.Application
$objExcel.visible = $True
$objWorkbook = $objExcel.Workbooks.Add()
$objSheet = $objWorkbook.Worksheets.Item(1)
$objSheet.Cells.Item(1,1) = "Server"
$objSheet.Cells.Item(1,2) = "LogName"
$objSheet.Cells.Item(1,3) = "Time"
$objSheet.Cells.Item(1,4) = "Source"
$objSheet.Cells.Item(1,5) = "Message"
$objSheetFormat = $objSheet.UsedRange
$objSheetFormat.Interior.ColorIndex = 19
$objSheetFormat.Font.ColorIndex = 11
$objSheetFormat.Font.Bold = $True

$row = 1

$servers = gc servers.txt

foreach ($server in $servers)
{
  $row = $row + 1
  $AppLog = Get-EventLog -LogName Application -EntryType Error -computer $server -Newest 5
  $SecLog = Get-EventLog -LogName Security -EntryType Error -computer $server -Newest 5 -ea Silentlycontinue
  $SysLog = Get-EventLog -LogName System -EntryType Error -computer $server -Newest 5
  foreach ($Cat in $AppLog,$Syslog,$Seclog)
  {
    if ($cat -is [array])
    {
      if ($AppLog -contains $cat[0]) {$Catname = "Application"}
      if ($SecLog -contains $cat[0]) {$Catname = "Security"}
      if ($SysLog -contains $cat[0]) {$Catname = "System"}
      Foreach ($event in $cat)
      {
        $objSheet.Cells.Item($row,1).Font.Bold = $True
        $objSheet.Cells.Item($row,1) = $server
        $objSheet.Cells.Item($row,2) = $Catname
        $objSheet.Cells.Item($row,3) = $Event.TimeGenerated
        $objSheet.Cells.Item($row,4) = $Event.Source
        $objSheet.Cells.Item($row,5) = $Event.Message
        $row = $row + 1
      }
    }
  }
}

$objSheetFormat = $objSheet.UsedRange
$objSheetFormat.EntireColumn.AutoFit()
$objSheetFormat.RowHeight = 15