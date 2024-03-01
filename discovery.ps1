# Discovery Script

# Command to get system information
systeminfo

# Command to list installed programs
Get-WmiObject -Query "SELECT * FROM Win32_Product" | Select-Object Name, Version

# Command to list network connections
Get-NetTCPConnection

# Command to list running processes
Get-Process

# Command to list services
Get-Service

# Command to list scheduled tasks
Get-ScheduledTask
