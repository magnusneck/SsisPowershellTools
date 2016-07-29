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

        # Private functions
       
        function shouldProcess {

            Param(
                [string[]]$types,
                [string]$componentType
            )

            ($types -contains $componentType -or $types -contains "All")
        }


        function writeTaskSql2008 {

            Param(
                [Parameter(Position=1)]
                [Xml]$xml,

                [Parameter(Position=2)]
                [Xml.XmlNamespaceManager]$ns,

                [Parameter(Position=3)]
                [string]$fileName,

                [Parameter(Position=4)]
                [string]$description   
            )

           
                $xml.SelectNodes("//DTS:Executable/DTS:Property[@DTS:Name='Description' and text()='$description']/..", $ns) | ForEach-Object {   
                $taskname = $_.SelectSingleNode("DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
                $sql = $_.SelectSingleNode("DTS:ObjectData/SQLTask:SqlTaskData/@SQLTask:SqlStatementSource", $ns).Value.ToString()
                writeSqlObject $fileName $taskname $description $taskname $sql
            }
        }


        function writeTaskSql2014 {

            Param(
                [Parameter(Position=1)]
                [Xml]$xml,

                [Parameter(Position=2)]
                [Xml.XmlNamespaceManager]$ns,

                [Parameter(Position=3)]
                [string]$fileName,

                [Parameter(Position=4)]
                [string]$description
            )

            $xml.SelectNodes("//DTS:Executable[@DTS:Description='$description']", $ns) | ForEach-Object {   
                $taskname = $_.GetAttribute("DTS:ObjectName")
                $sql = $_.SelectSingleNode("DTS:ObjectData/SQLTask:SqlTaskData/@SQLTask:SqlStatementSource", $ns).Value.ToString()
                writeSqlObject $fileName $taskname $description $taskname $sql
            }
        }


        function writeVariableSql2008 {

            Param(
                [Parameter(Position=1)]
                [Xml]$xml,

                [Parameter(Position=2)]
                [Xml.XmlNamespaceManager]$ns,

                [Parameter(Position=3)]
                [string]$fileName,

                [Parameter(Position=4)]
                [string]$description
            )

           $xml.SelectNodes("//DTS:Variable/DTS:Property[@DTS:Name='Expression' and text() != '']/..", $ns) | ForEach-Object {   
                $variableName = $_.SelectSingleNode("DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
                $expression = $_.SelectSingleNode("DTS:Property[@DTS:Name='Expression']", $ns).InnerText
                $namespace = $_.SelectSingleNode("DTS:Property[@DTS:Name='Namespace']", $ns).InnerText
                writeSqlObject $fileName $namespace $description $variableName $expression
            }
        }


        function writeVariableSql2014 {

            Param(
                [Parameter(Position=1)]
                [Xml]$xml,

                [Parameter(Position=2)]
                [Xml.XmlNamespaceManager]$ns,

                [Parameter(Position=3)]
                [string]$fileName,

                [Parameter(Position=4)]
                [string]$description
            )

           $xml.SelectNodes("//DTS:Variables/DTS:Variable[@DTS:Namespace != 'System']", $ns) | ForEach-Object {   
                $variableName = $_.GetAttribute("DTS:ObjectName")
                $expression = $_.GetAttribute("DTS:Expression")
                $namespace = $_.GetAttribute("DTS:Namespace")
                if ($expression -ne "") {
                    writeSqlObject $fileName $namespace $description $variableName $expression
                }
            }
        }


        function writeComponentSql2008 {

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
                [string]$fileName
            )

            $xpath = "//component[@description='$descriptionCondition']"
            $xml.SelectNodes($xpath, $ns) | ForEach-Object {   
                $taskname = $_.SelectSingleNode("../../../../DTS:Property[@DTS:Name='ObjectName']", $ns).InnerText
                $sql = $_.SelectSingleNode("properties/property[@name='$sqlAttributeName']", $ns).InnerText 
                $componentName = $_.SelectSingleNode("@name").Value.ToString()
                writeSqlObject $fileName $taskname $componentDescription $componentName $sql
            }
        }


        function writeComponentSql2014 {

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
                [string]$fileName
            )

            # TODO: Update xpath for 2014
            $xpath = "//component[@description='$descriptionCondition']"
            $xml.SelectNodes($xpath, $ns) | ForEach-Object {
                $taskname = $_.SelectSingleNode("../../../..").GetAttribute("DTS:ObjectName")
                $sql = $_.SelectSingleNode("properties/property[@name='$sqlAttributeName']", $ns).InnerText
                $componentName = $_.GetAttribute("name")
                writeSqlObject $fileName $taskname $componentDescription $componentName $sql
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


        if ($Path -like "*.dtsx") {
            Write-Verbose "File: $Path"

            [xml]$xml = Get-Content -Path $Path

            $ns = new-object Xml.XmlNamespaceManager ($xml.NameTable)
            $ns.AddNamespace("SQLTask", "www.microsoft.com/sqlserver/dts/tasks/sqltask")
            $ns.AddNamespace("DTS", "www.microsoft.com/SqlServer/Dts")
       
            $PackageFormatVersion = $xml.SelectSingleNode("/DTS:Executable/DTS:Property[@DTS:Name='PackageFormatVersion']", $ns).InnerText

            if ($PackageFormatVersion -notin ("3", "8")) {
                Write-Verbose "Warning: $Path has unsupported PackageFormatVersion: ${PackageFormatVersion}. Ignoring file."
            }
            else {           
                $description = "Execute SQL Task"
                if (shouldProcess $Type $description) {
                    switch ($PackageFormatVersion) {
                        "3" {writeTaskSql2008 $xml $ns $Path[0] $description}
                        "8" {writeTaskSql2014 $xml $ns $Path[0] $description}
                        default: {}
                    }
                }

                $description = "Variable"
                if (shouldProcess $Type $description) {
                    switch ($PackageFormatVersion) {
                        "3" {writeVariableSql2008 $xml $ns $Path[0] $description}
                        "8" {writeVariableSql2014 $xml $ns $Path[0] $description}
                        default: {}
                    }
                }

                $components = @(
                    @{descriptionCondition="OLE DB Source";  componentDescription="OLE DB Source"; sqlAttributeName="SqlCommand"}, 
                    @{descriptionCondition="OLE DB Destination"; componentDescription="OLE DB Destination"; sqlAttributeName="OpenRowset"},
                    @{descriptionCondition="Executes an SQL command for each row in a dataset."; componentDescription="OLE DB Command"; sqlAttributeName="SqlCommand"},
                    @{descriptionCondition="Looks up values in a reference dataset by using exact matching."; componentDescription="Lookup"; sqlAttributeName="SqlCommand"}
                )
               
                ForEach($component in $components) {
                    if (shouldProcess $Type $component.componentDescription) {
                        switch ($PackageFormatVersion) {
                            "3" {writeComponentSql2008 $xml $ns $component.descriptionCondition $component.componentDescription $component.sqlAttributeName $Path[0]}
                            "8" {writeComponentSql2014 $xml $ns $component.descriptionCondition $component.componentDescription $component.sqlAttributeName $Path[0]}
                            default: {}
                        }
                    }
                }
                   
                # TODO DFT:
                # - ADO NET Destination
                # - ADO NET Source
                # - SQL Server Destination
                #
                # TODO Control Flow:
                # - Bulk Insert Task

            }
        }
        else {
            Write-Verbose "File $Path is not a dtsx file"
        }
    }

    End
    {
    }
}
