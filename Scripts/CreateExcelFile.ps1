$a = New-Object -comobject Excel.Application

$a.Visible = $True

$b = $a.Workbooks.Add()
$c = $b.Worksheets.Item(1)

$c.Cells.Item(1,1) = 'This is Cell A1'
$b.SaveAs('..\Documents\GitHub\excel-automation\Spreadsheets\Excel PowerShell Test.xlsx')

$a.Quit()