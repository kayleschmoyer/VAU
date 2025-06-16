# VastAutoUpdater

VastAutoUpdater is a Windows Forms application written in VB.NET that checks a remote SFTP server for the latest VAST update and installs it when available. After each run it sends a summary email reporting success or failure.

## Project Structure

- **VastAutoUpdater/** – main application project
  - **Services/** – helper classes for SFTP, configuration, logging, email, and the update engine
  - **UI/** – Windows Forms UI implemented using MaterialSkin
  - **App.config** – application configuration (SFTP/SMTP settings)

The repository also contains a legacy copy of the project under `VastAutoUpdater/VastAutoUpdater` which is kept for reference.

## Building

Open `VastAutoUpdater.sln` in Visual Studio 2017 or later and build the solution. Credentials for email and SFTP are loaded from `App.config`.

## Usage

Run the compiled executable. Provide SFTP credentials when prompted or start with the `silent` argument to run without the UI. After completion a summary email will be sent detailing the outcome and timestamp.
