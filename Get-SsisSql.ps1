<#
.Synopsis
   Extract SQL queries from SSIS (dtsx) files.
.DESCRIPTION
   Extracts SQL queries from SQL Server Integration services (dtsx) files.
.EXAMPLE
   Get-SsisSql -Path MyFile.dtsx 

   Retrives all SQL queries from the MyFiles.dtsx file.
.EXAMPLE
   Get-ChildItem C:\Work\Release -Recurse | Get-SsisSql -Type Lookup, "OLE DB Source" | Where-Object { $_.SQL -like "*SomeTableName*" } 

   Find all references in LookUps and OLE DB Source components to "SomeTableName" in the SQL.
.OUTPUTS
SsisTools.SsisSql
- FileName = Name of the file
- TaskName = Name of the task in the Control flow
- ComponentName = Name of the component in the Data Flow Task. TaskName for Control flow tasks
- ComponentType = Type of component, e.g "OLE DB Source"
- SQL = The SQL query

.NOTES
The script currently supports SSIS 2008 file format.
#>
function Get-SsisSql
{
    [CmdletBinding(SupportsShouldProcess=$false, 
                  PositionalBinding=$false)]
     
    [OutputType('SsisTools.SsisSql')]

    Param
    (
        <#
          File names to search for SQL. Only *.dtsx files are queried, others are ignored. 
          This parameter can be Pipelined to the script, e.g by Get-ChildItem.
        #>
        [Parameter(Mandatory, 
                   ValueFromPipeline,
                   ValueFromPipelineByPropertyName, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [ValidateScript({Test-path $_})]
        [Alias("FileName", "FullName")] 
        [String[]]$Path,

        <#
        { "All", "Execute SQL Task", "OLE DB Source", "OLE DB Destination", "OLE DB Command", "Lookup", "Variable" }
   
        What kind of component types should be queried?
        #>
        [ValidateSet("All", "Execute SQL Task", "OLE DB Source", "OLE DB Destination", "OLE DB Command", "Lookup", "Variable")]
        [string[]]$Type = "All"
    )

    Begin
    {
    }

    Process
    {
        if ($Path -like "*.dtsx") {
            Write-Verbose "File: $Path"

            [xml]$xml = Get-Content -Path $Path

            $ns = new-object Xml.XmlNamespaceManager ($xml.NameTable)
            $ns.AddNamespace("SQLTask", "www.microsoft.com/sqlserver/dts/tasks/sqltask")
            $ns.AddNamespace("DTS", "www.microsoft.com/SqlServer/Dts")
        
            writeTaskSql $xml $ns $Path[0] $Type
            writeVariableSql $xml $ns $Path[0] $Type
            writeComponentSql $xml $ns 'OLE DB Source' 'OLE DB Source' 'SqlCommand' $Path[0] $Type
            writeComponentSql $xml $ns 'OLE DB Destination' 'OLE DB Destination' 'OpenRowset' $Path[0] $Type
            writeComponentSql $xml $ns 'Executes an SQL command for each row in a dataset.' 'OLE DB Command' 'SqlCommand' $Path[0] $Type
            writeComponentSql $xml $ns 'Looks up values in a reference dataset by using exact matching.' 'Lookup' 'SqlCommand' $Path[0] $Type

            # TODO DFT:
            # - ADO NET Destination
            # - ADO NET Source
            # - SQL Server Destination
            # 
            # TODO Control Flow:
            # - Bulk Insert Task

        }
        else {
            Write-Verbose "File $Path is not a dtsx file"
        }
    }

    End
    {
    }
}

function shouldProcess {

    Param(
        [string[]]$type,
        [string]$componentType
    )

    ($type -contains $componentType -or $type -contains "All")
}


function writeTaskSql {

    Param(
        [Parameter(Position=1)]
        [Xml]$xml,

        [Parameter(Position=2)]
        [Xml.XmlNamespaceManager]$ns,

        [Parameter(Position=3)]
        [string]$fileName,

        [Parameter(Position=4)]
        [string[]]$type
    )

    $description = "Execute SQL Task"
    if (shouldProcess $type $description) {
        $xml.SelectNodes("//DTS:Executable/DTS:Property[@DTS:Name='Description' and text()='$description']/..", $ns) | ForEach-Object {    
            $taskname = $_.SelectSingleNode("DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
            $sql = $_.SelectSingleNode("DTS:ObjectData/SQLTask:SqlTaskData/@SQLTask:SqlStatementSource", $ns).Value.ToString()
            writeSqlObject $fileName $taskname $description $taskname $sql
        }
    }

}


function writeVariableSql {

    Param(
        [Parameter(Position=1)]
        [Xml]$xml,

        [Parameter(Position=2)]
        [Xml.XmlNamespaceManager]$ns,

        [Parameter(Position=3)]
        [string]$fileName,

        [Parameter(Position=4)]
        [string[]]$type
    )

    if (shouldProcess $type "Variable") {
        $xml.SelectNodes("//DTS:Variable/DTS:Property[@DTS:Name='Expression' and text() != '']/..", $ns) | ForEach-Object {    
            $variableName = $_.SelectSingleNode("DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
            $expression = $_.SelectSingleNode("DTS:Property[@DTS:Name='Expression']", $ns).InnerText
            $namespace = $_.SelectSingleNode("DTS:Property[@DTS:Name='Namespace']", $ns).InnerText
            writeSqlObject $fileName $namespace 'Variable' $variableName $expression
        }
    }

}


function writeComponentSql {

    Param(
        [Parameter(Position=1)]
        [Xml]$xml,

        [Parameter(Position=2)]
        [Xml.XmlNamespaceManager]$ns,

        [Parameter(Position=3)]
        [string]$descriptionCondition,

        [Parameter(Position=4)]
        [string]$componentDescription,

        [Parameter(Position=5)]
        [string]$sqlAttributeName,

        [Parameter(Position=6)]
        [string]$fileName,

        [Parameter(Position=7)]
        [string[]]$type
    )

    if (shouldProcess $type $componentDescription) {
        $xpath = "//component[@description='$descriptionCondition']"
        $xml.SelectNodes($xpath, $ns) | ForEach-Object {    
            $taskname = $_.SelectSingleNode("../../../../DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
            $sql = $_.SelectSingleNode("properties/property[@name='$sqlAttributeName']", $ns).InnerText
            $componentName = $_.SelectSingleNode("@name").Value.ToString()
            writeSqlObject $fileName $taskname $componentDescription $componentName $sql
        }

    }
}


function writeSqlObject
{
    Param(
        [Parameter(Position=1)]
        [string]$FileName,
        
        [Parameter(Position=2)]
        [string]$TaskName,
        
        [Parameter(Position=3)]
        [string]$ComponentType,
        
        [Parameter(Position=4)]
        [string]$ComponentName,
        
        [Parameter(Position=5)]
        [string]$Sql
    )


    $prop = @{
        'FileName'=$FileName
        'TaskName'=$TaskName
        'ComponentName' = $ComponentName
        'ComponentType' = $ComponentType
        'SQL'=$Sql
    }
    $obj=New-Object -TypeName PSObject -Property $prop
    $obj.PSObject.TypeNames.Insert(0,’SsisTools.SsisSql’)
    Write-Output $obj
}
