# Get all Env variables for user
$envVar = [Environment]::GetEnvironmentVariables("User").Path

# Get current location
$location = Get-Location;

# Add location with 
$envVar = $envVar + $location

# Set new Env variables for user
[System.Environment]::SetEnvironmentVariable('Path', $EnvVar, [System.EnvironmentVariableTarget]::User)

# Install package for node.js
npm install
