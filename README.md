Logistics Dispatch Management System
==================
INTERFACE - PYTHON + ARDUINO + RFID Via USB

Dispatch Management is a process of managing the outgoing and incoming consignment packages intend to transport goods. The type of management system depends up nature of the products being sent. Automobile Harware
is a large sector that depends on Logistical Dispatch systems, where Parts of heavy weights are sent over crates of logistics. In bussiness, the dispatch system accounts for all the products packed in a particular
crate/Bin with reference to the invoices and relative to the total orders and quantity being sent over a particular bin/crate. This is where Dispatch management system comes to play.
This is an open source software i developed based on Python with PyQT5 library for UI, in windows. The software is integrated of two parts: Managining Software interface in Windows, Arduino Based RFID Interface connected
via USB.

Current Features
----------------
* Uses Excel Sheet based MasterData, Invoice Data, And Bin Data to input entries.
* RFID based Consignment Scanning Mapped via BIN SHEET.
* Dashboard Via User Login.
* Creates Shipment Via Measuring Total Weight of the Crate and Feeding data into RFID Card.
* Shipment Verification using Weigh Scale Interface via R232, through Invoice mapping provided with Weigh Error Tolerance.
* History of Verified Shipment Via JSON.
* Crate Label Generation After Verification with PDF.

Modules
----------------
* Main.py - Main execution module.
* RFID.py - Connects with Arduino coupled with RFID MFRC-522 - Any Arduino can be used. Pin Out provided in official MFRC 522 Documentation
* Weight.py - Connects to Weight Scale For Shipment Verification and Shipment Creation Connected via R232 Serial Port.
* ui.py - main GUI Interface Module for PyQT5
* Arduino Code - Contains C Script for Arduino With MFRC522 Module. Takes in MFRC Library provided in GITHUB.


