function which ($command) {
	Get-Command -Name $command -ErrorAction SilentlyContinue | 
		Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
}

Import-Module -Name Security -Force

function Out-Clipboard {
	[cmdletbinding()]
	param (
		[parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
		$InputObject,
        [switch] $File
	)
	begin {
		$ps = [PowerShell]::Create()
		$rs = [RunSpaceFactory]::CreateRunspace()
		$rs.ApartmentState = "STA"
		$rs.ThreadOptions = "ReuseThread"
		$rs.Open()
		$data = @()
	}
	process {
		$data += $InputObject
	}
	end {
        $rs.SessionStateProxy.SetVariable("do_file_copy", $File)
		$rs.SessionStateProxy.SetVariable("data", $data)
		$ps.Runspace = $rs
		$ps.AddScript({
			Add-Type -AssemblyName 'System.Windows.Forms'
            if ($do_file_copy) {
                 $file_list = New-Object -TypeName System.Collections.Specialized.StringCollection
                 $data | % {
                    if ($_ -is [System.IO.FileInfo]) {[void]$file_list.Add($_.FullName)} 
                    elseif ([IO.File]::Exists($_))   {[void]$file_list.Add($_)}
                }
                [System.Windows.Forms.Clipboard]::SetFileDropList($file_list)
            } else {
                $tmp_file = [IO.Path]::GetTempFileName()
                $data | Out-File -FilePath $tmp_file -Encoding Unicode
                $formatted_data = Get-Content -Path $tmp_file -Encoding Unicode
                Remove-Item -Path $tmp_file -ErrorAction SilentlyContinue -Force
			    [System.Windows.Forms.Clipboard]::SetText(($formatted_data -join "`n"))
            }
		}).Invoke()
	}
}