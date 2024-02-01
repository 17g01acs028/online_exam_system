# Service Monitor System

## Overview
The Service Monitor System is a Java-based application designed to check the status of various services as configured. It can monitor both server and application up/down times, ensuring real-time service availability.

## Features
- **Service Status Monitoring**: Monitors the status of services (up or down) at specific intervals.
- **Multiple Configuration File Support**: Supports various configuration files such as CSV, XML, INI, YAML, and JSON.
- **Log File Monitoring**: Tracks service status using log files, which are generated in the background once monitoring starts.
- **Log File Archiving**: Periodically archives log files to maximize space efficiency, based on specified archiving intervals.

## How It Works
1. **Service Configuration**: Reads service configuration details from files (CSV, XML, INI, YAML, JSON).
2. **Status Monitoring**: Monitors each service at specified intervals to check its status.
3. **Logging**: If enabled, logs service status in the background.
4. **Archiving**: Archives old log files at specified intervals if archiving is enabled.

## Setup and Configuration
1. **Clone/Download the Repository**: Obtain the code.
2. **Configuration Directory**: Place your service configuration files in the `config` directory.
3. **Running the Application**: Execute the `Main` class to start monitoring. The system reads the configuration file and starts monitoring services as per the configurations.

## Supported Configuration Files
- **CSV (.csv)**
- **XML (.xml)**
- **INI (.ini)**
- **YAML (.yaml)**
- **JSON (.json)**

Each configuration file should contain details about the services to monitor, such as service name, host, port, and monitoring intervals.

## Command Interface
The application offers a command-line interface to interact with the system. Available commands:
- `sky-monitor start`: Starts the monitoring process.
- `sky-monitor stop`: Stops the monitoring process.
- `sky-monitor application status <ID>`: Checks the status of a specific application by its ID.
- `sky-monitor server status <ID>`: Checks the status of a specific server by its ID.
- `sky-monitor service list`: Lists all services currently being monitored.
- `sky-monitor help`: Displays help information with a list of commands.
- `exit`: Exits the program.

## Dependencies
- Apache Commons CSV
- Google Gson
- SnakeYAML
- ini4j
- json-simple

Make sure these dependencies are included in your project to ensure smooth operation.

## Author
Stephen Mutio
