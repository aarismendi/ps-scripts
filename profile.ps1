function which ($command) {
	Get-Command -Name $command -ErrorAction SilentlyContinue | 
		Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
}

Import-Module -Name Security -Force

function Out-Clipboard {
	[cmdletbinding()]
	param (
		[parameter(ValueFromPipeline=$true)]
		$InputObject
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
		$rs.SessionStateProxy.SetVariable("data", $data)
		$ps.Runspace = $rs
		$ps.AddScript({
			Add-Type -AssemblyName 'System.Windows.Forms'
            $tmp_file = [IO.Path]::GetTempFileName()
            $data | Out-File -FilePath $tmp_file -Encoding Unicode
            $formatted_data = Get-Content -Path $tmp_file -Encoding Unicode
            Remove-Item -Path $tmp_file -ErrorAction SilentlyContinue -Force
			[System.Windows.Forms.Clipboard]::SetText(($formatted_data -join "`n"))
		}).Invoke()
	}
}