# Define parameters
param (
    [Parameter(Mandatory = $true)]
    [string]$RemoteComputer,  # Remote computer name or IP address

    [Parameter(Mandatory = $true)]
    [string]$AppVPackagePath, # Path to the App-V package (.appv file) on the remote computer

    [Parameter(Mandatory = $true)]
    [string]$AppVPackageName  # Name of the App-V package
)

# Function to publish App-V package
function Publish-AppVPackage {
    param (
        [string]$ComputerName,
        [string]$PackagePath,
        [string]$PackageName
    )

    # Establish a remote session
    Write-Host "Connecting to $ComputerName..." -ForegroundColor Cyan
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param ($Path, $Name)

        # Import the App-V module
        Import-Module AppvClient -ErrorAction Stop

        # Add the App-V package
        Write-Host "Adding App-V package: $Path" -ForegroundColor Yellow
        Add-AppvClientPackage -Path $Path | Out-Null

        # Publish the App-V package globally
        Write-Host "Publishing App-V package: $Name" -ForegroundColor Yellow
        Publish-AppvClientPackage -Name $Name -Global | Out-Null

        Write-Host "App-V package published successfully!" -ForegroundColor Green
    } -ArgumentList $PackagePath, $PackageName -ErrorAction Stop
}

# Call the function
Publish-AppVPackage -ComputerName $RemoteComputer -PackagePath $AppVPackagePath -PackageName $AppVPackageName
