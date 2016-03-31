#This function scrapes the TechNet page for Exchange Server build numbers and release
#dates to match the build numbers for Exchange Servers in the organization.
##Reference: Lee Holmes article on extracting tables from web pages was very useful for developing this
#Link: #http://www.leeholmes.com/blog/2015/01/05/extracting-tables-from-powershells-invoke-webrequest/
Function Get-ExchangeBuildNumbers()
{
    [CmdletBinding()]
    param ()
    
    Write-Verbose "Fetching Exchange build numbers from TechNet"
    
    $URL = "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.160).aspx"
    try
    {
        $WebPage = Invoke-WebRequest -Uri $URL -ErrorAction STOP
        $tables = @($WebPage.Parsedhtml.getElementsByTagName("TABLE"))

        Write-Verbose "Parsing results from web request"
        foreach ($table in $tables)
        {
            $rows = @($table.Rows)

            foreach($row in $rows)
            {
                $cells = @($row.Cells)

                ## If we’ve found a table header, remember its titles
                if($cells[0].tagName -eq "TH")
                {
                    $titles = @($cells | ForEach-Object { ("" + $_.InnerText).Trim() })
                    continue
                }

                ## If we haven’t found any table headers, make up names "P1", "P2", etc.
                if(-not $titles)
                {
                    $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
                }

                ## Now go through the cells in the the row. For each, try to find the
                ## title that represents that column and create a hashtable mapping those
                ## titles to content

                $resultObject = [Ordered] @{}

                for($counter = 0; $counter -lt $cells.Count; $counter++)
                {
                    $title = $titles[$counter]
                    if(-not $title) { continue }
                    $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()
                }

                ## And finally cast that hashtable to a PSCustomObject
                [PSCustomObject] $resultObject
            }
    }
    }
    catch
    {
        Write-Warning $_.Exception.Message
        $ExchangeBuildNumbers = "An error occurred. $($_.Exception.Message)"
    }

    return $ExchangeBuildNumbers
}

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$BuildNumbers = Get-ExchangeBuildNumbers

$BuildNumbers

$BuildNumbers | Export-Clixml $myDir\ExchangeBuildNumbers.xml
