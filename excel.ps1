$pwd = Convert-Path .
echo $pwd

$datas = @()

$path = $pwd + "\fileList.txt"
$array = Get-Content -Encoding UTF8 $path

#Initializing varriables
$excel = New-Object -ComObject Excel.Application
$book = $null

try {
    foreach($str in $array) {
        echo $str
        $path = $pwd + "\file\"
        $path += $str

        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        #open excel's book
        $book = $excel.Workbooks.Open($path)
        $sheet = $excel.Worksheets.Item("sheet1")

        $textLine = ""
        $line = 4

        #Repeat as long as the No. is not blank
        while($true) {
            $number = "B" + $line
            if([string]::IsNullOrEmpty($sheet.Range($number).Text)) {
                break
            }
            $textLine = $str + "|"

            #Repeat only the specified column
            foreach($column in "B", "C", "D") {
                $range = $column + $line
                $textLine += $sheet.Range($range).Text + "|"
            }
            $datas += $textLine
            $textLine = ""
            $line++
        }
        echo $datas.Length

    } 
    } finally {
        $excel.Quit()
        $excel = $null
        [GC]::Collect()
    }

#Initializing varriables
$excel = New-Object -ComObject Excel.Application
$book = $null

#Create a list
try {
    $path = $pwd + "\list.xlsx"
    echo $path

    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    #open excel's book
    $book = $excel.Workbooks.Open($path)
    $sheet = $excel.Worksheets.Item("sheet1")

    $line = 4
    foreach($data in $datas ) {
        $splitStr = $data.Split('|')
        $count = 0
        foreach($column in "A", "B", "C", "D") {
            $range = $column + $line
            $sheet.Range($range) = $splitStr[$count]
            $count++
        }
        $line++
    }
    $book.Save()
    echo Save Complite

} finally {
    $excel.Quit()
    $excel = $null
    [GC]::Collect()
}
