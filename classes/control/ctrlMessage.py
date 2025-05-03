import tkinter as tk
from tkinter import messagebox


class ctrlMessage():
    def view_top_warning(result):
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messages = ctrlMessage.get_message(result)
        messagebox.showwarning(messages[0], messages[1])
        return

    def view_top_okcancel(result):
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.withdraw()
        root.lift()
        root.focus_force()
        messages = ctrlMessage.get_message(result)
        return messagebox.askokcancel(messages[0], messages[1])

    def print_error(err):
        print(' --- Error message --- ')
        print(f" type   :[{str(type(err))}]")
        print(f" args   :[{str(err.args)}]")
        print(f" message:[{err.message}]")
        print(f" error  :[{err}]")

    def get_message(result):
        list_message = []
        messages = []
        list_message.append('[Code：')
        list_message.append(str(result))
        list_message.append(']\r\n')
        # formPackageMain
        if result == 0:
            list_message.append('　Process completed successfully.')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9999:
            list_message.append('　 ---- ')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9000:
            list_message.append('　Processing. Please wait for a moment...')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == 9001:
            list_message.append('　Do you want to proceed?')
            messages.append('Execution Confirmation')
            messages.append(''.join(list_message))
        elif result == 9002:
            list_message.append('　Process has been canceled.')
            messages.append('')
            messages.append(''.join(list_message))
        elif result == -1001:
            list_message.append('　An error occurred while switching the check mode.')
            messages.append('Mode Switch Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1002:
            list_message.append('　An error occurred while reverting to packaging mode.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1101:
            list_message.append('　The required field "Target Folder" is empty.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1102:
            list_message.append('　The specified folder for "Target Folder" does not exist.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1111:
            list_message.append('　The required field "Work Target" is empty.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1121:
            list_message.append('　The required field "Work Date" is empty or contains invalid values.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1131:
            list_message.append('　The required field "Department Name / Worker Name" is empty.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1141:
            list_message.append('　The required field "Work Terminal Name" is empty.')
            messages.append('Required Field Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1201:
            list_message.append('　The total file size in the target folder exceeds the threshold.')
            messages.append('File Count and Size Check Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1202:
            list_message.append('　The number of files in the target folder exceeds the threshold.')
            messages.append('File Count and Size Check Error｜formPackageMain')
            messages.append(''.join(list_message))
        elif result == -1301:
            list_message.append('　An error occurred while copying report files from the installation folder.')
            messages.append('Report File Copy Error｜formPackageMain')
            messages.append(''.join(list_message))
        # ctrlCsv
        elif result == -3001:
            list_message.append('　An error occurred while outputting the CSV file.')
            messages.append('CSV Output Error｜ctrlCsv')
            messages.append(''.join(list_message))
        elif result == -3101:
            list_message.append('　An error occurred while converting Manifest.xml to a CSV intermediate file.')
            messages.append('XML to CSV Conversion Error｜ctrlCsv')
            messages.append(''.join(list_message))
        elif result == -3201:
            list_message.append('　An error occurred while converting multiple Manifest.xml files to CSV intermediate files.')
            messages.append('Multiple XML to CSV Conversion Error｜ctrlCsv')
            messages.append(''.join(list_message))
        # ctrlBatch
        elif result == -4101:
            list_message.append('　An exceptional error occurred while retrieving the PowerShell version. (Version information is not numeric)')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4102:
            list_message.append('　The PowerShell version is less than 7.')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4111:
            list_message.append('　The pwsh command was not found. PowerShell 7 or later is not installed, or the system has not been restarted after installation. Alternatively, review the environment variable definitions.')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4112:
            list_message.append('　An error occurred during the pwsh version check.')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4201:
            list_message.append('　Please execute in a Windows environment.')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4202:
            list_message.append('　An error occurred during PowerShell execution.')
            messages.append('Batch Processing Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4301:
            list_message.append('　The storage location cannot be opened because the input content on the screen has been changed after execution.')
            messages.append('Storage Location Opening Error｜ctrlBatch')
            messages.append(''.join(list_message))
        elif result == -4302:
            list_message.append('　An error occurred while opening the storage location.')
            messages.append('Storage Location Opening Error｜ctrlBatch')
            messages.append(''.join(list_message))
        # ctrlExcel
        elif result == -5001:
            list_message.append('　Excel from Office 2019 or later is not installed.')
            messages.append('Microsoft Excel Error｜ctrlExcel')
            messages.append(''.join(list_message))
        elif result == -5002:
            list_message.append('　The Excel application cannot be referenced. Please install Office 2019 or later.')
            messages.append('Microsoft Excel Error｜ctrlExcel')
            messages.append(''.join(list_message))
        # CommonModules.psm1
        elif result == -6001:
            list_message.append('　This PowerShell must be executed with version 7.0 or later.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6101:
            list_message.append('　An error occurred during the folder deletion process.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6102:
            list_message.append('　An error occurred during the file deletion process.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6201:
            list_message.append('　An error occurred during the compression process to a ZIP file.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6301:
            list_message.append('　An error occurred during the extraction process of a ZIP file.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6401:
            list_message.append('　An error occurred while creating a new folder during the folder recreation process. Location: New-DirectoryIfNotExists')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6402:
            list_message.append('　An error occurred while recreating a folder during the folder recreation process. Location: New-DirectoryIfNotExists')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        elif result == -6501:
            list_message.append('　A discrepancy occurred when comparing data before and after packaging.')
            messages.append('PowerShell Script Error｜CommonModules.psm1')
            messages.append(''.join(list_message))
        # AdpackController.ps1
        elif result == -7001:
            list_message.append('　Required external module files (*.ps1, *.psm1) are missing.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7002:
            list_message.append('　An error occurred while loading external module files (*.ps1, *.psm1).')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7003:
            list_message.append('　Please review the arguments. [Reason: Both Pack + UnPack are set]')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7004:
            list_message.append('　Please review the arguments. [Reason: Neither Pack nor UnPack is set]')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7005:
            list_message.append('　Please review the arguments. [Reason: Pack + Check, or Pack + UnCheck is set]')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7006:
            list_message.append('　Please review the arguments. [Reason: UnPack + Check + UnCheck all three are set]')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7007:
            list_message.append('　Please specify "Folder" for input data.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7008:
            list_message.append('　Please specify "File (*.zip)" for output data in Pack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7009:
            list_message.append('　Please specify "File (*.zip)" for input data in UnPack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7010:
            list_message.append('　Please specify "Folder" for output data in UnPack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7011:
            list_message.append('　The executable file for Pack does not exist.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        # AdpackModules.psm1
        elif result == -7101:
            list_message.append('　An error occurred while creating a folder to store XML files.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7102:
            list_message.append('　An error occurred while creating Index.xml.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7103:
            list_message.append('　An error occurred while creating Manifest.xml.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7104:
            list_message.append('　A mismatch occurred during the hash check.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7105:
            list_message.append('　An error occurred during the packaging process using Pack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7106:
            list_message.append('　An exceptional error occurred during the packaging process using Pack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7107:
            list_message.append('　An error occurred during the check process using Pack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7108:
            list_message.append('　An error occurred during the unpack process using Pack (NoCheck argument included).')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7109:
            list_message.append('　An error occurred while deleting decompressed data after the check process using Pack.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        elif result == -7110:
            list_message.append('　An error occurred while renaming the temporary folder during the custom module unpack process.')
            messages.append('PowerShell Script Error｜AdpackController.ps1')
            messages.append(''.join(list_message))
        # MultiCheckController.ps1
        elif result == -7201:
            list_message.append('　Required external module files (*.ps1, *.psm1) are missing.')
            messages.append('PowerShell Script Error｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7202:
            list_message.append('　An error occurred while loading external module files (*.ps1, *.psm1).')
            messages.append('PowerShell Script Error｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7203:
            list_message.append('　Please specify "Folder" for input data.')
            messages.append('PowerShell Script Error｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7204:
            list_message.append('　The executable file for Pack does not exist.')
            messages.append('PowerShell Script Error｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        elif result == -7205:
            list_message.append('　No ZIP files exist in the specified folder.')
            messages.append('PowerShell Script Error｜MultiCheckController.ps1')
            messages.append(''.join(list_message))
        # PrintController.ps1
        elif result == -8003:
            list_message.append('　Please review the arguments. [Reason: Both FormA + FormB are set]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8004:
            list_message.append('　Please review the arguments. [Reason: Neither FormA nor FormB is set]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8005:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: RootPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8006:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: TemplatePath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8007:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: formPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8008:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: headerMappingPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8009:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: bodyMappingPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8010:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: headerValuesPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8011:
            list_message.append('　The folder or file specified in the arguments does not exist. [Target: bodyValuesPath]')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8012:
            list_message.append('　The default sheet was not found in the form template file (Excel).')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8013:
            list_message.append('　An error occurred during the copy process of the form template file (Excel).')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8014:
            list_message.append('　An error occurred while reading position data and input data for header information.')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        elif result == -8015:
            list_message.append('　The item names in position data and input data for header information do not match.')
            messages.append('PowerShell Script Error｜PrintController.ps1')
            messages.append(''.join(list_message))
        # Screen-related error messages
        elif result == -9001:
            list_message.append('　Multiple instances are prohibited.')
            messages.append('Warning: Multiple Instances Prohibited')
            messages.append(''.join(list_message))
        else:
            list_message.append('　An exception occurred, and the process terminated abnormally.')
            messages.append('')
            messages.append(''.join(list_message))
        return messages
