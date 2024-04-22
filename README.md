# PrinterManagement

## 🚀 Description
This PowerShell script facilitates printer management in Windows environments, offering functionalities such as installing and removing printers, sending bulk test pages, cleaning print queues, and conducting an inventory of printers connected to the system.

## 📋 Contents
1. **Printer Installation:** Allows the installation of printers specified in a CSV file containing details such as the name, local port, and driver.
   
2. **Printer Deletion:** Facilitates the deletion of printers and their associated ports based on information provided in a CSV file.
   
3. **Test Page Sending:** Sends test pages to printers specified in a CSV file to verify their proper functionality.
   
4. **Print Queue Cleaning:** Stops the print queue service, deletes all pending jobs, and restarts the service. It is recommended not to use during productive hours.
   
5. **Printer Inventory:** Generates a CSV file containing detailed information about printers installed on the system, including the name, port, driver, and status.

## ⚙️ Usage
1. Clone or download the repository.
2. Run the `PrinterManagement.ps1` script from PowerShell.
3. Create a CSV file with the next format:
4. ```powershell
    Name;LocalPort;DriverName
    printer1;\\computer\printer1;Epson
    printer2;192.168.1.156;Epson
    ```
5. Select the desired option from the menu.

## 📝 Requirements
- PowerShell 5.1 or higher.
- Administrator permissions to perform certain actions.

## ⚠️ Notes
- This script must be run with administrator privileges.
- Cleaning the print queue (option 4) may interrupt the printing of ongoing documents.
