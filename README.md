# ICL-Server-Connector
A Python-based GUI application for configuring and executing Broadband Forum telecom standards like TR114i2 and TR115i3. Designed for performance testing of DSLAM and DUT setups, the tool enables users to input test setup details,select from list of standardized tests,execute them via an active ICL server running on the host machine.


# Overview
**ICL Server Connector** is a Python-based GUI application designed to interface with an ICL server for executing standardized telecom test procedures defined by the Broadband Forum, such as **TR114i2**, **TR115i3**, and others. The tool facilitates performance validation of **DSLAM** (Digital Subscriber Line Access Multiplexer) and **DUT** (Device Under Test) setups by enabling users to configure test environments, select test cases, and execute them in real-time.


# Features
- GUI built with **Tkinter** and **TTK** for intuitive user interaction
- **Multi-threaded architecture** for concurrent socket communication and UI responsiveness
- **Socket-based client-server communication** with the ICL server
- Dynamic test setup configuration and execution
- Automated logging and result tracking
- Execution of ICL language functions for telecom standards
- Integration with Excel-based test data using **Openpyxl**
- Robust error handling and configuration management
- Modular design for extensibility and hardware abstraction

# Architecture
[User GUI] → [Threaded Command Dispatcher] → [Socket Client] → [ICL Server] ↑ ↓ [Config Parser] ← [Excel Test Data] ← [Test Selection Module]
- **GUI Layer**: Built with Tkinter/TTK for user input and test control
- **Threaded Dispatcher**: Manages command queues and server responses
- **Socket Client**: Communicates with the ICL server over TCP/IP
- **Test Modules**: Encapsulate logic for executing TR114i2, TR115i3, etc.
- **Data Layer**: Reads setup parameters and test data from Excel files

# Supported Standards
- TR114i2 – Performance Testing for DSL Access
- TR115i3 – Functional Testing for Broadband Devices
- Additional Broadband Forum standards as required

# Supported Devices
- DMS4
- DMS12
- SFPG5321
- SFPV5311
- BCM_UDP
- Other DUTs compatible with ICL language execution

# Installation

# Prerequisites
- Python 3.8+
- ICL Server running on host machine
- Excel files with test setup data


# Packaging the Application (Windows Executable)
pyinstaller --onefile --windowed --add-data "sparnex-logo.jpg;." --hidden-import=tkinter --hidden-import=openpyxl --hidden-import=pandas --hidden-import=ttkbootstrap --collect-all ttkbootstrap ServerConnector_Stable.py  

# Setup
git clone https://github.com/paddu-tumuluri/ICL-Server-Connector.git
cd ICL-Server-Connector
python main.py

# Author
Padmini Tumuluri
R&D Engineer | Python Developer | QA Automation Specialist
LinkedIn Profile https://www.linkedin.com/in/padmini-tumuluri-963133121/
