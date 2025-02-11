# Revit Cloud Model Opener

## Overview
Revit Cloud Model Opener is a PowerShell script designed to streamline the process of opening Revit models hosted on BIM 360. The script provides a user-friendly GUI that allows users to filter, select, and open Revit models with various options such as creating new local files, detaching from central, loading worksets, and performing audits. The script also supports running custom Dynamo scripts on the models.

## Requirements
- Revit Batch Processor must be installed.
- PowerShell must be installed on your system.
- .NET Framework (required for Windows Forms)

![image](https://github.com/user-attachments/assets/b3eb4c9b-c2e0-46c9-950f-3a24a407e6be)


## Features
- **Model Management**:
  - Open Revit models from BIM 360 with ease.
  - Options to create new local files or detach from central.
  - Load or close all worksets.
  - Perform audits on models.
- **Script Execution**:
  - Run custom Dynamo scripts on selected models.
  - Use default minimal Python script if no custom script is selected.
- **User-Friendly GUI**:
  - Filter models by project name, folder path, and file name.
  - Adjust font size of the list view for better visibility.
  - Refresh the list of models based on the selected filters.
- **Performance Optimization**:
  - Toggle fast mode to keep Revit open between operations for faster performance.

## Installation
1. Ensure that Revit Batch Processor is installed on your system.
2. Save the `RevitCloudOpener.ps1` script to a desired location on your computer.

## Usage
1. Open PowerShell and navigate to the directory where `RevitCloudOpener.ps1` is saved.
2. Run the script using the command:
   ```powershell
   .\RevitCloudOpener.ps1
