# createExcelInventoryReport.ps1

Foobar is a Python library for dealing with word pluralization.
createExcelInventoryReport.ps1 is a small inventory script for VMware, to generate a Excel report.

## Installation

To use this script the Powershell modules VMware.PowerCLI and ImportExcel need to be installed. Both are available from PSGallery.

```powershell
Install-Module -Name VMware.PowerCLI
Install-Module -Name ImportExcel
```

## Usage

```powershell
.\createExcelInventoryReport.ps1
```

An example of the Excel report can be found [here](Example-vCenterClusterInfo-withVM-20200728.xlsx)

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

N/A