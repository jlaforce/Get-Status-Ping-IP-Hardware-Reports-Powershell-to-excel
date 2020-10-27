$path = ".\hostname-results.xls" 
 
$objExcel = new-object -comobject excel.application  
 
if (Test-Path $path)  
{  
$objWorkbook = $objExcel.WorkBooks.Open($path)  
$objWorksheet = $objWorkbook.Worksheets.Item(1)  
} 
 
else {  
$objWorkbook = $objExcel.Workbooks.Add()  
$objWorksheet = $objWorkbook.Worksheets.Item(1) 
} 
 
$objExcel.Visible = $True 
 
#########Add Header#### 
 
$objWorksheet.Cells.Item(1, 1) = "MachineIP" 
$objWorksheet.Cells.Item(1, 2) = "Result" 
$objWorksheet.Cells.Item(1, 3) = "HostName" 
 
$machines = Get-Content .\pcs.txt 
$count = $machines.count 
 
$row=2 
 
$machines | foreach-object{ 
$ping=$null 
$hname =$null 
$machine = $_ 
$ping = Test-Connection $machine -Count 1 -ea silentlycontinue 
 
if($ping){ 
 
$objWorksheet.Cells.Item($row,1) = $machine 
$objWorksheet.Cells.Item($row,2) = "UP" 
     
$hname = [System.Net.Dns]::GetHostByAddress($machine).HostName 
 
$objWorksheet.Cells.Item($row,3) = $hname  
         
$row++} 
else { 
 
$objWorksheet.Cells.Item($row,1) = $machine 
$objWorksheet.Cells.Item($row,2) = "DOWN" 
 
$row++} 
} 