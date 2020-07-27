# Need PowerCLI module and ImportExcel
#Requires –Modules VMware.PowerCLI
#Requires –Modules ImportExcel
# Install-Module -Name VMware.PowerCLI
# Install-Module -Name ImportExcel

$xlfile = "C:\temp\vCenterClusterInfo-withVM-"+ $enddate +".xlsx"

Connect-VIServer -Server x.x.x.x, x.x.x.x

$report = Foreach($vc in $global:DefaultVIServers){
  foreach($dc in Get-Datacenter -Server $vc){
    Get-Cluster -Location $dc -Server $vc -PipelineVariable cluster |
      Get-VMHost |
        Select @{N='VC';E={$vc.Name}},
        @{N='Datacenter';E={$dc.Name}},
        @{N='Cluster';E={$cluster.Name}},
        @{N='#ESXi';E={$cluster.ExtensionData.Host.Count}},
        Name,
        @{N='ESXi version';E={"$($_.Version) $($_.Build)"}},
        @{N='ESXi HW';E={"$($_.ExtensionData.Hardware.SystemInfo.Vendor) $($_.ExtensionData.Hardware.SystemInfo.Model)"}},
        @{N='ESXi CPU';E={$_.ProcessorType}},
        @{N='ESXi CPU Count';E={$_.ExtensionData.Hardware.CpuInfo.NumCpuPackages}},
        @{N='ESXi Cores';E={$_.ExtensionData.Hardware.CpuInfo.NumCpuCores}},
        @{N='ESXi Hyperthreading';E={$_.HyperthreadingActive}},
        @{N='ESXi PowerState';E={$_.PowerState}},
        @{N='ESXi ConnectionState';E={$_.ConnectionState}}
  }
}
$report2 = Foreach($vc in $global:DefaultVIServers){
  foreach($dc in Get-Datacenter -Server $vc){
    Get-Cluster -Location $dc -Server $vc -PipelineVariable cluster |
      Get-VMHost -PipelineVariable vmhost |
      Get-VM |
        Select @{N='VC';E={$vc.Name}},
        @{N='Datacenter';E={$dc.Name}},
        @{N='Cluster';E={$cluster.Name}},
        @{N='ESX Host';E={$vmhost.Name}},
        Name,
        PowerState,
        ResourcePool,
        Guest,
        GuestId,
        NumCpu,
        MemoryMB,
        Folder,
        ExtensionData.Guest.GuestFullName
  }
}
$enddate = (Get-Date).tostring("yyyyMMdd")
Remove-Item $xlfile -ErrorAction SilentlyContinue
$excel = Open-ExcelPackage -Path "$xlfile" -Create
$excel = $report | Export-Excel -ExcelPackage $excel -AutoSize -WorkSheetname "vCenter Clusters Info" -AutoFilter -tablename vCenter -tablestyle Medium2 -PassThru
Add-PivotTable -ExcelPackage $excel -PivotTableName "Cluster Summary" -SourceRange $excel.Workbook.Worksheets[1].Tables[0].Address -PivotRows Datacenter,Cluster -PivotData @{"ESXi Cores"="sum";"ESXi CPU Count"="sum";"Name"="count"} -PivotDataToColumn -PivotTotals "None" -NoTotalsInPivot -Activate
$excel = $report2 | Export-Excel -ExcelPackage $excel -AutoSize -WorkSheetname "VM Info" -AutoFilter -tablename VM -tablestyle Medium2 -PassThru
Add-PivotTable -ExcelPackage $excel -PivotTableName "VM Summary" -SourceRange $excel.Workbook.Worksheets[3].Tables[0].Address -PivotRows Datacenter,Cluster,ResourcePool -PivotData @{"Name"="count"} -PivotFilter PowerState,GuestId -PivotDataToColumn -PivotTotals "None" -NoTotalsInPivot -Activate
Close-ExcelPackage $excel -Show