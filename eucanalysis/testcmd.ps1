# Example pipeline cmdlet

function test {
  param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)] $value
  )

    begin {
        $i=0
        write-host "Begin"
    }
  
    process {
        $i++
        write-host "Process Item " $i $value.id $value.filepath 
    
    }

    end {
        write-host "End"
    }
      
}


#Example 1
#$global:j=1
#dir|select @{Name="id";Expression={$global:j; $global:j++}}, @{Name="filepath";Expression={$_.name}}|test


#Example 2
dir|% {$i=0} {@{id=$i++;filepath=$_.name}}|test
