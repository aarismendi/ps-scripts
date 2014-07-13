function which ($command) {
	Get-Command -Name $command -ErrorAction SilentlyContinue | 
		Select-Object -ExpandProperty Path -ErrorAction SilentlyContinue
}

Import-Module -Name Security -Force

function Unzip-File {
    [cmdletbinding(DefaultParameterSetName="ByPath", SupportsShouldProcess=$True)]
    param (
        [parameter(Mandatory=$true, ParameterSetName="ByPath", Position=0)] [string] $Path,
        [parameter(Mandatory=$true, ParameterSetName="ByInput", ValueFromPipeline=$true)] $InputObject,
        [string] $Destination,
        [switch] $Force
    )
    begin {$shell_app = new-object -com shell.application}
    process {
        switch ($PSCmdlet.ParameterSetName) {
            'ByPath'  {$input_paths = Resolve-Path -Path $Path | ? {[IO.File]::Exists($_.Path)} | Select -Exp Path}
            'ByInput' {if ($InputObject -is [System.IO.FileInfo]) {$input_paths = $InputObject.FullName}}
        }
        $input_paths | % {   
            $zip_file = $shell_app.namespace($_)
            if (-not $Destination) {
                $dest_home = (Get-Item -Path $_).Directory.FullName
            } else {
                $dest_home = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)
            }
            $dest_name = [IO.Path]::GetFileNameWithoutExtension($_)
            $zip_dest = [IO.Path]::Combine($dest_home, $dest_name)

            if (-not (Test-Path -Path $zip_dest -PathType Container)) {
                New-Item -Path $zip_dest -ItemType Directory -Force | Out-Null
            }
            $dest = $shell_app.namespace($zip_dest)
            $NODIALOG = 0x4
            $YESTOALL = 0x10
            if ($Force) {
                $dest.Copyhere($zip_file.items(), ($YESTOALL -bor $NODIALOG))
            } else {
                $dest.Copyhere($zip_file.items(), $NODIALOG)
            }
        }
    }
}

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
<<<<<<< HEAD
            if ($do_file_copy) {
                 $file_list = New-Object -TypeName System.Collections.Specialized.StringCollection
                 $data | ForEach-Object {
                    if ($_ -is [System.IO.FileInfo]) {[void]$file_list.Add($_.FullName)} 
                    elseif ([IO.File]::Exists($_))   {[void]$file_list.Add($_)}
                }
                [System.Windows.Forms.Clipboard]::SetFileDropList($file_list)
            } else {
                $sdata = ($data | Out-String) -split "`n"
			    [System.Windows.Forms.Clipboard]::SetText($sdata)
            }
=======
			if ($do_file_copy) {
				$file_list = New-Object -TypeName System.Collections.Specialized.StringCollection
				$data | % {
					if ($_ -is [System.IO.FileInfo]) {[void]$file_list.Add($_.FullName)} 
					elseif ([IO.File]::Exists($_))   {[void]$file_list.Add($_)}
				}
						[System.Windows.Forms.Clipboard]::SetFileDropList($file_list)
			} else {
				$non_printable = '[\x20\x00\x08\x0B\x0C\x0E-\x1F]+?$'
				$host_out = (($data | Out-String -Width 2000) -split "`n" | % 
					{$_.TrimEnd() -replace $non_printable, ''}) -join "`n"
				[System.Windows.Forms.Clipboard]::SetText($host_out)
			}
		}).Invoke()
	}
}

function Select-Folder {
	[cmdletbinding()]
	param (
		[string] $Message = "Select Directory:", [System.Environment+SpecialFolder] $RootFolder
	)
	begin {}
	process {
		$ps = [PowerShell]::Create()
		$rs = [RunSpaceFactory]::CreateRunspace()
		$rs.ApartmentState = "STA"
		$rs.ThreadOptions = "ReuseThread"
		$rs.Open()
		$sel_path = $null
		$rs.SessionStateProxy.SetVariable("sel_path", $sel_path)
		$rs.SessionStateProxy.SetVariable("Message", $Message)
		$rs.SessionStateProxy.SetVariable("RootFolder", $RootFolder)
		$ps.Runspace = $rs
		$ps.AddScript({
			Add-Type -AssemblyName System.Windows.Forms
			$folder_browser = New-Object -TypeName Windows.Forms.FolderBrowserDialog
			$folder_browser.Description = $Message
			$folder_browser.RootFolder = $RootFolder
			$outer = New-Object -TypeName System.Windows.Forms.Form
			$outer.StartPosition = [Windows.Forms.FormStartPosition] "Manual"
			$outer.Location = New-Object System.Drawing.Point -500, -500
			$outer.Size = New-Object System.Drawing.Size 0, 0
			$outer.add_Shown({ 
				$outer.Activate()
				$result = $folder_browser.ShowDialog($outer)
				$outer.Close()
			})
			$outer.ShowDialog() | Out-Null
			if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
				$sel_path = $folder_browser.SelectedPath
			}
>>>>>>> 8f5bcc22c1ccbc6942e054c22146cb22ab5b0599
		}).Invoke()
		$sel_path = $rs.SessionStateProxy.GetVariable("sel_path")
		if ($sel_path) {return $sel_path}
	}
	end {}
}