[CmdletBinding()]
PARAM(
	[Parameter(Mandatory=$True)]
	[string[]]
	$vCenter,
	
	[Parameter(Mandatory=$True)]
	[ValidateScript({Test-Path $_})]
	[string]
	$CsvPath
)

BEGIN {
	Import-Module Vmware.PowerCLI
	if ( $Global:DefaultViServer ) {
		Disconnect-ViServer * -Force -Confirm:$false
	}
}

PROCESS {
	ForEach ( $viServer in $vCenter ) {
		Connect-ViServer $viServer
		$TagData = Import-Csv $CsvPath
		
		$Properties = $TagData[0].PSObject.Properties.Name
		$Categories = $Properties | Where-Object { $_.Split(".")[0] -eq "Category" }
		
		ForEach ( $Category in $Categories ) {
			$CategoryName = $Category.Split(".")[2].Trim()
			$Cardinality = $Category.Split(".")[1]
			
			$EntityTypes = $TagData | Where-Object { -not([string]::IsNullOrEmpty($_.$($Category))) }  | Select-Object -ExpandProperty "Entity.Type" | Sort-Object -Unique
			
			$ExistingCategory = Get-TagCategory -Name $CategoryName -ErrorAction SilentlyContinue
			
			if ( $null -eq $ExistingCategory) {
				Write-Verbose "$CategoryName : Creating category  as $Cardinality"
				$ExistingCategory = New-TagCategory -Name $CategoryName -Cardinality $Cardinality -EntityType $EntityTypes
			}
			else {
				Write-Verbose "$CategoryName : already exists."
				[array]$NewEntityTypes = $EntityTypes | Where-Object { [array]($ExistingCategory.EntityType) -notcontains $_ }
				if ($NewEntityTypes.Count -gt 0 ) {
					Write-Verbose "$CategoryName : Adding new entity types $NewEntityTypes"
					$ExistingCategory | Set-TagCategory -AddEntityType $NewEntityTypes
				}
			}
			
			$Tags = $TagData.$($Category) | Sort-Object -Unique | Where-Object { -NOT([string]::isNullOrEmpty($_)) }
			
			ForEach ( $TagF in $Tags ) {
				$Tag = $TagF.Trim()
				Write-Verbose "$CategoryName : Tag: '$Tag'"
				
				$TagObj = Get-Tag -Name $Tag -Category $CategoryName -ErrorAction SilentlyContinue
				if ($null -eq $TagObj) {
					Write-Verbose "$CategoryName : creating tag $tag"
					$TagObj = New-Tag -Name $Tag -Category $CategoryName
					$TagObj = Get-Tag -Name $Tag -Category $CategoryName -ErrorAction SilentlyContinue
				}
				
				[array]$VmsToTag = $TagData | Where-Object { $_.$($Category) -eq $Tag } | Select-Object -ExpandProperty "Entity.Name"
				[array]$Vms = Get-Vm $VmsToTag -ErrorAction SilentlyContinue
				$msg = "{0} VMs ({1} found) with tag {2} for category {3}" -f $VmsToTag.Count, $Vms.Count, $Tag, $CategoryName
				Write-Verbose $msg
				if ( $Vms.Count -gt 0 ) {
					$Vms | New-TagAssignment -Tag $TagObj | Out-Null
				}
			}
		}
		
		Disconnect-ViServer $viServer -Force -Confirm:$False
	}
}
