#A way to turn CSVs from clients into update scripts faster, 
#while avoiding a few gotchas involving quotation marks and NULL values

$csv_path = './test.csv' #Put your path to the CSV here
$my_table = Import-Csv -Path $csv_path
$headers = $my_table | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
$header_count = $headers.Count

#This grabs the first row of the CSV to get the header names.
#They must match the column names of the table it is being inserted into
$insert_into = 'INSERT INTO table_name ('
foreach ($header in $headers) {
    $insert_into = $insert_into + $header + ','
}
$insert_into = ($insert_into + ') VALUES').Replace(',)', ')')

#This grabs the second row of the CSV which should just be 1's and 0's
#A value of 1 indicates you want to wrap that column in quotation marks,
#and anything else indicates that you don't.
$add_quotes = @()
for ($i = 0; $i -lt $header_count; $i++) {
    $add_quotes += ($my_table[0].($headers[$i]))
}

for ($i = 1; $i -lt $my_table.Count; $i++) {
    #On every 5000th line (including line 0), put the Insert Into statement generated
    #above into the query. SQL Server can only handle 5000 lines at a time.
    if (($i-1)%5000 -eq 0) {
        Write-Host $insert_into
    }

    $row = $my_table[$i] #Just a helper assignment for readability
    
    $print_string = '('
    for ($j = 0; $j -lt $header_count; $j++) {
        if ($add_quotes[$j] -eq 1) {
            #Add starting quote mark if requested
            $print_string = $print_string + "'"
        }

        #Get string information if you can
        $print_string = $print_string + $row.($headers[$j])

        if ($add_quotes[$j] -eq 1) {
            $print_string = $print_string + "'"
        }

        if ($j + 1 -ne $header_count) {
            #Add closing comma if not last item in row
            $print_string = $print_string + ', '
        }
    }
    #Add closing parentheses and comma. Not a good way to prevent final comma though
    $print_string = $print_string + '),'
    
    #Replace any 'NULL's with NULLs
    $print_string = $print_string.Replace("''NULL''", 'NULL').Replace("'NULL'", 'NULL')
    Write-Host $print_string
}
