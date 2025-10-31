# Laser Hypot Continuity

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Usage](#usage)
4. [Technical](#technical)
5. [Syntax](#syntax)
6. [Contributions](#contributions)
7. [Installation](#installation)

## Overview

This project is designed to read output and send to a shared network drive from a Vapor Pro XL Moisture Analyzer

## Features

- Reads Data via USB B serial connection
- Writes and formats data to an XLSX file

## Usage

- Double click .exe to start the program, any time data is "printed" or when a test ends-
  - It will create a local copy of the document
  - Then it copies the file based on cached credentials to the designated UNC network path

## Technical

- Python 3.14
- **Moisture Analyzer Model:** Computrac Vapor Pro XL <https://www.brookfieldengineering.com/products/moisture-analyzer/computrac-vapor-pro-xl>
- Developed and tested on Windows 10 Pro

## Syntax

- **Variables**: snake_case
- **Functions**: snake_case

## Contributions

- Made, created, and designed by Tony Martin at Matrix Plastic Products

## Installation

- It is recommended to disable "Allow the computer to turn off this device to save power" in control panel on the USB Hubs, 
or else they can lose connection and need to be unplugged and replugged<br>

- To install this project, clone the repository and install the required dependencies:

  - Install USB Serial Driver <https://www.brookfieldengineering.com/products/moisture-analyzer/-/media/227dd0ba7bfb4d839fcbe9ac8bd750c2.ashx?dmc=1&la=en&revision=1f3de233-3b6e-4f06-9878-84a39d38e85a>

```sh
git clone https://github.com/matrixplastic/Vapor-Pro-XL-Data-Output.git
cd /Path/To/Your/Cloned/Project
pip install -r requirements.txt
```
## License

MIT License
