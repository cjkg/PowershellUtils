#A way to turn CSVs into update scripts faster, without excel,
#while avoiding a few gotchas involving quotation marks and NULL values

#Get CSV name, making sure that the file exists
do {
    $csv_name = Read-Host -Prompt 'Input the name of your CSV (include extension)'
    if ($csv_name -notlike '*.csv') {
        Write-Host "The name given does not have '.csv' as its extension. Try again."
    } elseif (!(Test-Path "./$csv_name")) {
        Write-Host "No CSV named $csv_name in the current folder with that name. Try again."
    } else {
        $csv_exists = $true
    }
} while (!($csv_exists))

#Get table name, making sure a null value or blank value is not given
do {
    $table_name = Read-Host -Prompt 'Input the table name you are writing the script for'
    if ($null -eq $table_name) {
        $tbl_name_length = 0
    } else {
        $tbl_name_length = ($table_name.Trim()).Length
    }
} while ($tbl_name_length -eq 0)

#Get sql file name, and make sure that it has the correct extension and isn't blank or empty
do {
    $out_name = Read-Host -Prompt 'Enter what name you want the output file to be named (include extension)'
    if ($null -eq $out_name) {  #need this case to avoid null error thrown
        $out_name_length = 0
    } elseif ($out_name -notlike '*.sql') { #just to make sure it's executable sql
        $out_name_length = 0
        Write-Host "The name given doesn't have a '.sql' extension. Try again."
    } else {
        $out_name_length = ($out_name.Trim()).Length #Otherwise get the length to break the loop
    }
} while ($out_name_length -eq 0)

$in_path = "./$csv_name"
$out_path = "./$out_name"

$my_table = Import-Csv -Path $in_path #import the csv

$headers = (Get-Content $in_path)[0] -Split ',' #Get csv headers. Could use Get-Member, but that puts them out of order
$header_count = $headers.Count

#This grabs the first row of the CSV to get the header names.
#They must match the column names of the table it is being inserted into
$insert_into = "INSERT INTO $table_name ("
foreach ($header in $headers) {
    $insert_into = $insert_into + $header + ','
}
$insert_into = "$insert_into) VALUES".Replace(',)', ')')

#This grabs the second row of the CSV which should just be 1's and 0's
#A value of 1 indicates you want to wrap that column in quotation marks,
#and anything else indicates that you don't.
$add_quotes = @()
for ($i = 0; $i -lt $header_count; $i++) {
    $add_quotes += ($my_table[0].($headers[$i]))
}

$final_statement = @()
for ($i = 1; $i -lt $my_table.Count; $i++) {
    #On every 5000th line (including line 0), put the Insert Into statement generated
    #above into the query. SQL Server can only handle 5000 lines at a time.
    if (($i-1)%5000 -eq 0) {
        $final_statement += $insert_into
    }

    $row = $my_table[$i] #Just a helper assignment for readability

    $print_string = '(' #start of the insert values list, opening parenthesis

    for ($j = 0; $j -lt $header_count; $j++) {
        if ($add_quotes[$j] -eq 1) { #if the options row = 1, add quotes, otherwise just add the value
            $print_string = $print_string + "'" + $row.($headers[$j]) + "'"
        } else { 
            $print_string = $print_string + $row.($headers[$j])
        }

        if ($j + 1 -ne $header_count) {
            #Add closing comma if not last item in values statement
            $print_string = $print_string + ', '
        }
    }
    #Add closing parentheses and comma (if applicable)
    if (($i+1 -ne $my_table.Count) -and ($i%5000 -ne 0)) {
        $print_string = $print_string + '),'
    } else {
        $print_string = $print_string + ')'
    }

    #Replace any 'NULL'/''NULL''s with NULLs to avoid them being inserted as strings
    $print_string = $print_string.Replace("''NULL''", 'NULL').Replace("'NULL'", 'NULL')
    $final_statement += $print_string
}

if (!(Test-Path $out_path)) { #If the outpath doesn't exist, create a file. Otherwise confirm first.
    Write-Host "File Created at $out_path" #Alert user
    New-Item -path $out_path #Create file
    Set-Content -Path $out_path -Value $final_statement #Add the INSERT INTO statements into the file
} else {
    do {
        $overwrite = Read-Host -Prompt "File already exists at $out_path. Overwrite? (Y/N)" #Confirm overwrite
    } while ($overwrite -notin @('N', 'Y'))

    if ($overwrite -eq 'Y') {
        Write-Host "File Overwritten at $out_path" #Alert User
        Set-Content -Path $out_path -Value $final_statement #Change the contents of the existing file
    } else {
        Write-Host "Aborting..." #Alert User
    }
}
