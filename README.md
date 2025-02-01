# **RTF to DOCX Converter**

## **Preface**
This project was developed to address the need for converting system-generated documents from **RTF** format to **DOCX**. Due to limitations in modifying the system’s default behavior, this script was created as a workaround to facilitate the conversion process seamlessly.

## **Overview**
The script runs continuously in the background, checking every second for the presence of a flag file (**txt**). When the flag file is detected, the script scans predefined directories where documents are created, converts **RTF** files to **DOCX**, and deletes the original **RTF** files.

### **Flag-Based Execution**
The flag mechanism is implemented to optimize system performance. The system workflow can include a command like:
```python
Create file=flag.txt
While flag.txt exists:
    sleep(1)
```
This ensures that the system can proceed with further processing once the conversion is complete, while the script itself remains lightweight and efficient.

Additionally, the script can be modified to accept arguments, such as a specific file path, for manual execution. If automatic execution on server startup is required, please refer to the **Notes** section.

## **Installation**
1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/yourrepository.git
   ```
2. Ensure **Microsoft Office** is installed on the system.
3. Install the required **pywin32** library:
   ```bash
   python -m pip install pywin32
   ```

## **Notes**
### **Running as a Windows Batch Job (Task Scheduler or Service)**
Running this script as a scheduled task or Windows service may present challenges due to Microsoft’s default restrictions on running Office applications in service mode. By default, the **SYSTEM** user lacks an environment for executing Microsoft Office applications like **Word**.

To bypass this limitation, manually create the necessary directories to allow Office automation:

- **For Windows Server (x64):**
  ```
  C:\Windows\SysWOW64\config\systemprofile\Desktop
  ```
- **For Windows Server (x86):**
  ```
  C:\Windows\System32\config\systemprofile\Desktop
  ```

If the issue persists, refer to the following resources:
- [Microsoft KB Article 257757](https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/257757)
- [MSDN Forum Discussion](https://social.msdn.microsoft.com/Forums/en-US/b81a3c4e-62db-488b-af06-44421818ef91/excel-2007-automation-on-top-of-a-windows-server-2008-x64?forum=innovateonoffice)

## **Support**
If you encounter any issues, have questions, or find bugs, feel free to reach out.
