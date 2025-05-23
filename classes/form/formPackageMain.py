import sys
import os
import shutil
from  pathlib import Path
import tkinter as tk
import tkinter.ttk as ttk
import tkinterdnd2 as tkdnd2
import tkcalendar as tkcal
from tkinter import messagebox
from tkinter import PhotoImage
from tkinter import filedialog
from datetime import datetime

# Load custom modules
from classes.structure import structureEntrydata
from classes.control import ctrlBatch
from classes.control import ctrlConfig
from classes.control import ctrlMessage
from classes.control import ctrlString
from classes.control import ctrlCsv
from classes.control import ctrlExcel
from classes.form import formBase
from classes.form.formAuth import formAuth

class formPackageMain(formBase):
    def __init__(self, master):
        # Call the parent class
        super().__init__(master)

        # Initial settings
        # _xmldata = ctrlConfig.get_xmldata('settings.xml')
        self.settings["installdir"] = 'C:/PyTkinterToPSScript'
        self.settings_filepath = self.settings['installdir'] + '/config/settings.xml'
        _xmldata = ctrlConfig.get_xmldata(self.settings_filepath)
        if _xmldata is None:
            messagebox.showerror('Error', 'The settings file does not exist or is empty.')
            sys.exit(-1)
        self.settings = self._add_settings(_xmldata)

        # Create objects
        self.master.title(self.settings.get('title'))
        self.master.geometry(self.settings.get('window_size'))
        self._create_header()
        self._create_body()
        self._create_footer()
        master.bind('<KeyPress>', self._key_handler)

    def __del__(self):
        print('close_process')

    def _add_settings(self, xmldata):
        # Parent class settings
        settings = self.settings

        # Read XML file
        settings = ctrlConfig.read_xmlfile(settings, xmldata)

        # ComboBox settings
        for combosettings in xmldata.findall('combosettings'):
            _array = []
            for _combodata in combosettings.findall('targetrange'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['targetrange'] = _array
            _array = []
            for _combodata in combosettings.findall('workername'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['workername'] = _array
            _array = []
            for _combodata in combosettings.findall('terminalname'):
                for _value in _combodata.findall('value'):
                    _array.append(_value.text)
                settings['terminalname'] = _array

        # Additional settings
        settings['title'] = 'Call PowerShell script from Python Tkinter'
        settings['label_logo'] = 'PyTkinterToPSScript'
        settings["label_result_init"] = ' ---- '
        settings['label_result'] = '        '
        settings['text_messages'] = 'No messages.'
        settings['label_mode_package'] = 'Current Mode｜Package'
        settings['label_mode_check'] =   'Current Mode｜Check'
        settings['button_modechange'] = 'Switch Mode'
        settings['label_basicsettings'] = 'Base Settings'
        settings['label_targetfolder'] = 'Target Folder'
        settings['textbox_folderpath'] = ''
        settings['button_databrose'] = 'Browse'
        settings['label_reportsettings'] = 'Report Settings'
        settings['label_targetrange'] = 'Target Range'
        settings['label_workdate'] = 'Work Date'
        settings['label_workername'] = 'Worker Name'
        settings['label_terminalname'] = 'Terminal Name'
        # Text color
        settings['fg_color'] = '#696969'
        settings['fg_color_package'] = '#ffffff'
        settings['fg_color_check'] = '#ffffff'
        # Background color
        settings['bg_color'] = '#f2f2f2'
        settings['bg_color_package'] = '#8b0000'
        settings['bg_color_check'] = '#006400'
        # Label settings
        settings['fb_color_ok'] = '#00008b'
        settings['fb_color_ng'] = '#ff0000'
        settings['fb_color_normal'] = '#000000'

        return settings

    # Key handler
    def _key_handler(self, event):
        if (event.keysym == 'F1' and
                self._button_f1['state'] == 'normal'):
            self._click_f1()
        elif (event.keysym == 'F2' and
                self._button_f2['state'] == 'normal'):
            self._click_f2()
        elif (event.keysym == 'F3' and
                self._button_f3['state'] == 'normal'):
            self._click_f3()
        elif (event.keysym == 'F4' and
                self._button_f4['state'] == 'normal'):
            self._click_f4()
        elif (event.keysym == 'F5' and
                self._button_f5['state'] == 'normal'):
            self._click_f5()
        elif (event.keysym == 'F6' and
                self._button_f6['state'] == 'normal'):
            self._click_f6()
        elif (event.keysym == 'F7' and
                self._button_f7['state'] == 'normal'):
            self._click_f7()
        elif (event.keysym == 'F8' and
                self._button_f8['state'] == 'normal'):
            self._click_f8()
        elif (event.keysym == 'F9' and
                self._button_f9['state'] == 'normal'):
            self._click_f9()
        elif (event.keysym == 'F10' and
                self._button_f10['state'] == 'normal'):
            self._click_f10()
        elif (event.keysym == 'F11' and
                self._button_f11['state'] == 'normal'):
            self._click_f11()
        elif (event.keysym == 'F12' and
                self._button_f12['state'] == 'normal'):
            self._click_f12()

    # Control definitions --->

    # Header area
    def _create_header(self):
        # Logo
        self._current_dir = os.path.dirname(__file__)
        self._image_path = os.path.join(
            self._current_dir,
            '..',
            'image', 'logo.png'
        )
        self._logoImage = PhotoImage(file=self._image_path)
        self._label_logo = tk.Label(
            self.frame_header,
            width=340,
            height=60,
            image=self._logoImage
        )
        self._label_logo.grid(row=0, column=0, sticky="w")
        # Processing result
        self._label_result = tk.Label(
            self.frame_header,
            text=self.settings['label_result'],
            font=self.settings['font_label_h2'],
            bg=self.settings['bg_color']
        )
        self._label_result.grid(row=0, column=1, sticky="e")
        # Message content
        _frame_messages = tk.Frame(
            self.frame_header,
            bg=self.settings['bg_color'],
            padx=0,
            pady=0
        )
        _frame_messages.grid(row=0, column=2, columnspan=3, sticky="e")

        self._text_messages = tk.Text(
            _frame_messages,
            width=55,
            height=3,
            wrap=tk.NONE,
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._text_messages.grid(row=0, column=0, sticky="nsew")

        # Apply scrollbars
        scrollbar_y = tk.Scrollbar(
            _frame_messages,
            command=self._text_messages.yview
        )
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x = tk.Scrollbar(
            _frame_messages,
            orient='horizontal',
            command=self._text_messages.xview
        )
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        # Apply scrollbars
        self._text_messages.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Initial value reflection
        self._text_messages.configure(state='normal')
        self._text_messages.delete('1.0', 'end')
        self._text_messages.insert('1.0', self.settings['text_messages'])
        self._text_messages.configure(state='disabled')
        # Mode
        self._label_mode = tk.Label(
            self.frame_header,
            text=self.settings['label_mode_package'],
            font=self.settings['font_label_h2'],
            fg=self.settings['fg_color_package'],
            bg=self.settings['bg_color_package']
        )
        self._label_mode.grid(row=1, column=3, sticky="e")
        # Mode change button
        self._button_modechange = tk.Button(
            self.frame_header,
            text=self.settings['button_modechange'],
            font=self.settings['font_button'],
            command=self._click_modechange)
        self._button_modechange.grid(row=1, column=4, sticky="e")
        # Layout
        self.frame_header.rowconfigure(0, weight=2)
        self.frame_header.rowconfigure(1, weight=1)
        self.frame_header.columnconfigure(0, weight=2)  # Logo
        self.frame_header.columnconfigure(1, weight=1)  # Processing result
        self.frame_header.columnconfigure(2, weight=20)  # Message (1/3)
        self.frame_header.columnconfigure(3, weight=1)  # Message (2/3) + Mode label
        self.frame_header.columnconfigure(4, weight=1)  # Message (3/3) + Mode button

    # Main area
    def _create_body(self):
        # Basic information
        self._frame_basicsettings = tk.Frame(
            self.frame_body,
            bg=self.settings['bg_color']
        )
        self._frame_basicsettings.pack(expand=1, fill="both", side="top")
        # Label
        self._label_basicsettings = tk.Label(
            self._frame_basicsettings,
            text=self.settings['label_basicsettings'],
            font=self.settings['font_label_h1'],
            bg=self.settings['bg_color']
        )
        self._label_basicsettings.grid(row=0, column=0, sticky="w")
        # Target folder - Label
        self._label_targetfolder = tk.Label(
            self._frame_basicsettings,
            text=self.settings['label_targetfolder'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_targetfolder.grid(row=1, column=0, sticky="w")
        # Target folder - Input box
        self._textbox_folderpath = tk.Entry(
            self._frame_basicsettings,
            textvariable=self.settings['textbox_folderpath'],
            font=self.settings['font_text']
        )
        self._textbox_folderpath.grid(row=1, column=1, sticky="ew")
        self._textbox_folderpath.drop_target_register(tkdnd2.DND_FILES)
        self._textbox_folderpath.dnd_bind('<<Drop>>', self.drop)
        # Browse button
        self._button_databrowse = tk.Button(
            self._frame_basicsettings,
            text=self.settings['button_databrose'],
            font=self.settings['font_button'],
            command=self._click_browse)
        self._button_databrowse.grid(row=1, column=2, sticky="ew")
        # Layout
        self._frame_basicsettings.columnconfigure(0, weight=1)
        self._frame_basicsettings.columnconfigure(1, weight=7)
        self._frame_basicsettings.columnconfigure(2, weight=1)

        # Report information
        self._frame_reportsettings = tk.Frame(
            self.frame_body,
            bg=self.settings['bg_color']
        )
        self._frame_reportsettings.pack(expand=1, fill="both", side="bottom")
        # Label
        self._label_reportsettings = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_reportsettings'],
            font=self.settings['font_label_h1'],
            bg=self.settings['bg_color']
        )
        self._label_reportsettings.grid(row=0, column=0, sticky="w")
        # Work target - Label
        self._label_targetrange = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_targetrange'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_targetrange.grid(row=1, column=0, sticky="w")
        # Work target - ComboBox
        self._combobox_targetrange = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['targetrange'],
            font=self.settings['font_text']
        )
        self._combobox_targetrange.grid(row=1, column=1, sticky="ew")
        # Work date - Label
        self._label_workdate = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_workdate'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_workdate.grid(row=2, column=0, sticky="w")
        # Work date - Calendar
        self._dateentry_workdate = tkcal.DateEntry(
            self._frame_reportsettings,
            font=self.settings['font_cal'],
            date_pattern='mm/dd/yyyy',
            firstweekday='sunday',
            showweeknumbers=False)
        self._dateentry_workdate.grid(row=2, column=1, sticky="w")
        self._dateentry_workdate.delete(0, tk.END)
        # Department name / Worker name - Label
        self._label_workername = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_workername'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_workername.grid(row=3, column=0, sticky="w")
        # Department name / Worker name - ComboBox
        self._combobox_workername = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['workername'],
            font=self.settings['font_text']
        )
        self._combobox_workername.grid(row=3, column=1, sticky="ew")
        # Work terminal name - Label
        self._label_terminalname = tk.Label(
            self._frame_reportsettings,
            text=self.settings['label_terminalname'],
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_terminalname.grid(row=4, column=0, sticky="w")
        # Work terminal name - ComboBox
        self._combobox_terminalname = ttk.Combobox(
            self._frame_reportsettings,
            state='normal',
            values=self.settings['terminalname'],
            font=self.settings['font_text']
        )
        self._combobox_terminalname.grid(row=4, column=1, sticky="ew")
        # Layout
        self._frame_reportsettings.columnconfigure(0, weight=1)
        self._frame_reportsettings.columnconfigure(1, weight=6)

    def _create_footer(self):
        # Create F1 button
        self._label_f1 = tk.Label(
            self.frame_footer, text='F1',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f1.grid(row=0, column=0, sticky="ew")
        self._button_f1 = tk.Button(
            self.frame_footer,
            text=' Package ',
            font=self.settings['font_button'],
            command=self._click_f1
        )
        self._button_f1.grid(row=1, column=0, sticky="ew")
        # Create F2 button
        self._label_f2 = tk.Label(
            self.frame_footer,
            text='F2',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f2.grid(row=0, column=1, sticky="ew")
        self._button_f2 = tk.Button(
            self.frame_footer,
            text=' Check ',
            font=self.settings['font_button'],
            command=self._click_f2
        )
        self._button_f2.grid(row=1, column=1, sticky="ew")
        self._button_f2.configure(state='disable')
        # Create F3 button
        self._label_f3 = tk.Label(
            self.frame_footer,
            text='F3',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f3.grid(row=0, column=2, sticky="ew")
        self._button_f3 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f3
        )
        self._button_f3.grid(row=1, column=2, sticky="ew")
        self._button_f3.configure(state='disable')
        # Create F4 button
        self._label_f4 = tk.Label(
            self.frame_footer,
            text='F4',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f4.grid(row=0, column=3, sticky="ew")
        self._button_f4 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f4
        )
        self._button_f4.grid(row=1, column=3, sticky="ew")
        self._button_f4.configure(state='disable')
        # Space
        self._label_space01 = tk.Label(
            self.frame_footer,
            text=' ',
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_space01.grid(row=0, column=4, rowspan=2, sticky="ew")
        # Create F5 button
        self._label_f5 = tk.Label(
            self.frame_footer,
            text='F5',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f5.grid(row=0, column=5, sticky="ew")
        self._button_f5 = tk.Button(
            self.frame_footer,
            text='Open Output Destination',
            font=self.settings['font_button'],
            command=self._click_f5
        )
        self._button_f5.grid(row=1, column=5, sticky="ew")
        self._button_f5.configure(state='disable')
        # Create F6 button
        self._label_f6 = tk.Label(
            self.frame_footer,
            text='F6',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f6.grid(row=0, column=6, sticky="ew")
        self._button_f6 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f6
        )
        self._button_f6.grid(row=1, column=6, sticky="ew")
        self._button_f6.configure(state='disable')
        # Create F7 button
        self._label_f7 = tk.Label(
            self.frame_footer,
            text='F7',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f7.grid(row=0, column=7, sticky="ew")
        self._button_f7 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f7
        )
        self._button_f7.grid(row=1, column=7, sticky="ew")
        self._button_f7.configure(state='disable')
        # Create F8 button
        self._label_f8 = tk.Label(
            self.frame_footer,
            text='F8',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f8.grid(row=0, column=8, sticky="ew")
        self._button_f8 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f8
        )
        self._button_f8.grid(row=1, column=8, sticky="ew")
        self._button_f8.configure(state='disable')
        # Space
        self._label_space02 = tk.Label(
            self.frame_footer,
            text=' ',
            font=self.settings['font_label_body'],
            bg=self.settings['bg_color']
        )
        self._label_space02.grid(row=0, column=9, rowspan=2, sticky="ew")
        # Create F9 button
        self._label_f9 = tk.Label(
            self.frame_footer,
            text='F9',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f9.grid(row=0, column=10, sticky="ew")
        self._button_f9 = tk.Button(
            self.frame_footer,
            text=' Settings ',
            font=self.settings['font_button'],
            command=self._click_f9
        )
        self._button_f9.grid(row=1, column=10, sticky="ew")
        self._button_f9.configure(state='disable')
        # Create Execute button
        self._label_f10 = tk.Label(
            self.frame_footer,
            text='F10',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f10.grid(row=0, column=11, sticky="ew")
        self._button_f10 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f10
        )
        self._button_f10.grid(row=1, column=11, sticky="ew")
        self._button_f10.configure(state='disable')
        # Create Settings button
        self._label_f11 = tk.Label(
            self.frame_footer,
            text='F11',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f11.grid(row=0, column=12, sticky="ew")
        self._button_f11 = tk.Button(
            self.frame_footer,
            text='      ',
            font=self.settings['font_button'],
            command=self._click_f11
        )
        self._button_f11.grid(row=1, column=12, sticky="ew")
        self._button_f11.configure(state='disable')
        # Create Close button
        self._label_f12 = tk.Label(
            self.frame_footer,
            text='F12',
            font=self.settings['font_label_body'],
            fg=self.settings['fg_color'],
            bg=self.settings['bg_color']
        )
        self._label_f12.grid(row=0, column=13, sticky="ew")
        self._button_f12 = tk.Button(
            self.frame_footer,
            text='Close',
            font=self.settings['font_button'],
            command=self._click_f12
        )
        self._button_f12.grid(row=1, column=13, sticky="ew")

        self.frame_footer.columnconfigure(0, weight=2)
        self.frame_footer.columnconfigure(1, weight=2)
        self.frame_footer.columnconfigure(2, weight=2)
        self.frame_footer.columnconfigure(3, weight=2)
        self.frame_footer.columnconfigure(4, weight=1)
        self.frame_footer.columnconfigure(5, weight=2)
        self.frame_footer.columnconfigure(6, weight=2)
        self.frame_footer.columnconfigure(7, weight=2)
        self.frame_footer.columnconfigure(8, weight=2)
        self.frame_footer.columnconfigure(9, weight=1)
        self.frame_footer.columnconfigure(10, weight=2)
        self.frame_footer.columnconfigure(11, weight=2)
        self.frame_footer.columnconfigure(12, weight=2)
        self.frame_footer.columnconfigure(13, weight=2)

    # Drop control
    def drop(self, event):
        self._textbox_folderpath.delete(0, tk.END)
        self._textbox_folderpath.insert(tk.END, event.data)

    # Button-related processing --->

    # Mode change button
    def _click_modechange(self):
        _result = 0

        # Read the password from the settings file
        correct_password = self.settings['password']
        # Switch to check mode with password authentication
        if self._button_f2["state"] == 'disabled':
            _password = formAuth(self.master, "Password Authentication").result
            # Cancelled during password authentication
            if _password is None:
                _result = 9999

            # Password authentication succeeded
            elif _password == correct_password:
                try:
                    # Switch to check mode (2)
                    self._change_mode(2)
                    
                except Exception as err:
                    _result = -1301
            # Password authentication failed
            else:
                tk.messagebox.showerror('Password Authentication Failed', 'Incorrect password.')
                _result = 9999

        # Switch back to package mode without authentication
        elif self._button_f2["state"] == 'normal':
            try:
                # Switch to package mode (1)
                self._change_mode(1)
            except Exception as err:
                _result = -1302

        # Do not output messages in normal cases
        if _result != 0:
            self._view_message(_result)
        
        return _result
    
    # Browse button
    def _click_browse(self):
        # Initial directory position
        if not os.path.isdir(self._textbox_folderpath.get()):
            currentdir = os.getcwd()
            initdir = currentdir
        else:
            initdir = os.path.dirname(self._textbox_folderpath.get())
        # Show dialog
        folder_path = filedialog.askdirectory(initialdir=initdir)
        if folder_path:
            self._textbox_folderpath.delete(0, tk.END)
            self._textbox_folderpath.insert(tk.END, folder_path)

    # F1 button
    def _click_f1(self):
        result = 0
        # Confirmation message before execution
        if not (ctrlMessage.view_top_okcancel(9001)) > 0:
            result = 9002
        # Message during execution
        if result == 0:
            self._view_message(9000)
        # Input check
        if result == 0:
            entry_data = self._get_entry_data()
            result = self._validate_entry(entry_data)
        # Output input information to CSV file
        if result == 0:
            output_path = self.settings['installdir'] + '/input/FormA_HeaderValues.csv'
            result = self._output_csv_headervalues(output_path)
        # Check file count and file size
        if result == 0:
            max_size = int(self.settings['package_maxsize'])        # Package file size 1.5GB = 1,610,612,736 = 1.5 * 1GB(1,073,741,824B)
            max_files = int(self.settings['package_maxfiles'])
            # Check with default file count and size
            target_path = entry_data.folderpath
            result = self._check_folder_limits(target_path, max_size, max_files)
        # Packaging
        if result == 0:
            # Specify the PowerShell script to execute
            script_path = self.settings['installdir'] + '/script/AdpackController.ps1'
            # Specify arguments for the PowerShell script
            hash_algorithm = self.settings['hash_algorithm']
            input_path = entry_data.folderpath
            output_path = '{}.zip'.format(input_path)
            output_path = self._get_unique_filename(output_path)
            # Execute packaging
            result, zip_path = self._packaging_data(script_path, hash_algorithm, input_path, output_path)
        # Output Manifest file to CSV after packaging
        if result == 0:
            input_path = entry_data.folderpath + '/META-INF/Manifest.xml'
            output_path = self.settings['installdir'] + '/input/FormA_BodyValues.csv'
            result = self._output_csv_formalistbodyvalues(input_path, output_path)
        # Output Form A list
        if result == 0:
            script_path = self.settings['installdir'] + '/script/PrintController.ps1'
            from_path = self.settings['installdir'] + '/output/FormALists.pdf'
            script_args = {
                'output_path': from_path,
                'root_path': self.settings['installdir'],
                'datamapping_header': 'DataMapping_Header.csv',
                'dataMapping_body': 'DataMapping_Body.csv',
                'forma_template': 'Template_FormALists.xlsx',
                'forma_headerValues': 'FormA_HeaderValues.csv',
                'forma_bodyValues': 'FormA_BodyValues.csv',
                'formb_template': 'Template_FormBLists.xlsx',
                'formb_headerValues': 'FormB_HeaderValues.csv',
                'formb_bodyValues': 'FormB_BodyValues.csv'
            }
            result = self._print_formalists(script_path, script_args)
        # Copy process (copy from install folder to target folder)
        if result == 0:
            # Specify download destination (copy destination)
            _target_folder = Path(entry_data.folderpath)
            to_path = _target_folder.parent.as_posix() + '/' + _target_folder.name + '.pdf'
            to_path = self._get_unique_filename(to_path)
            # Download process (copy)
            result, pdf_path = self._copy_formdata(from_path, to_path)
        # Enable "Open Output Destination" button if copy succeeds
        if result == 0:
            self._button_f5.configure(state='normal')
        # Output message
        self._view_message(result)
        # Append message for specific status codes
        list_message = []
        # Normal completion
        if result == 0:
            list_message.append('\r\n')
            list_message.append('　・ZIP file: ' + zip_path + '\r\n')
            list_message.append('　・PDF file: ' + pdf_path)
        # Append message
        self._append_message("".join(list_message))
    
    # F2 button
    def _click_f2(self):
        result = 0
        # Confirmation message before execution
        if not (ctrlMessage.view_top_okcancel(9001)) > 0:
            result = 9002
        # Change to "Executing"
        if result == 0:
            self._view_message(9000)
        # Input check
        if result == 0:
            entry_data = self._get_entry_data()
            result = self._validate_entry(entry_data)
        # Output input information to CSV file
        if result == 0:
            output_path = self.settings['installdir'] + '/input/FormB_HeaderValues.csv'
            result = self._output_csv_headervalues(output_path)
        # Check file count and file size
        if result == 0:
            target_path = entry_data.folderpath
            max_size = int(self.settings['check_maxsize'])
            max_files = int(self.settings['check_maxfiles'])
            result = self._check_folder_limits(target_path, max_size, max_files)
        # Check packaged data
        if result == 0:
            # Specify the PowerShell script to execute
            script_path = self.settings['installdir'] + '/script/MultiCheckController.ps1'
            # Specify arguments for the PowerShell script
            input_path = entry_data.folderpath
            output_path = self.settings['installdir'] + '/input/FormB_ZipFileList.csv'
            result = self._check_data(script_path, input_path, output_path)
        # Output Manifest file to CSV after packaging
        if result == 0:
            _zipfile_lists = ctrlCsv.read_csv(output_path)
            filepath_lists = []
            for _row in _zipfile_lists:
                _file_path = _row['FilePath']
                # Remove extension
                if _file_path.lower().endswith('.zip'):
                    _file_path = os.path.splitext(_file_path)[0]
                # Replace "\\" with "/"
                _file_path = _file_path.replace('\\', '/')

                filepath_lists.append(_file_path)

            output_path = self.settings['installdir'] + '/input/FormB_BodyValues.csv'
            result = self._output_csv_formblist_bodyvalues(filepath_lists, output_path)
        # Output Form A list
        if result == 0:
            script_path = self.settings['installdir'] + '/script/PrintController.ps1'
            script_args = {
                'output_path': self.settings['installdir'] + '/output/FormBLists.pdf',
                'root_path': self.settings['installdir'],
                'datamapping_header': 'DataMapping_Header.csv',
                'dataMapping_body': 'DataMapping_Body.csv',
                'formb_template': 'Template_FormBLists.xlsx',
                'formb_headerValues': 'FormB_HeaderValues.csv',
                'formb_bodyValues': 'FormB_BodyValues.csv'
            }
            result = self._print_formblists(script_path, script_args)
        # Download process
        if result == 0:
            # Specify data in the install folder
            from_path = self.settings['installdir'] + '/output/FormBLists.pdf'
            # Specify download destination (copy destination)
            _target_folder = Path(entry_data.folderpath)
            to_path = _target_folder.parent.as_posix() + '/' + _target_folder.name + '.pdf'
            to_path = self._get_unique_filename(to_path)
            result, pdf_path = self._copy_formdata(from_path, to_path)
        if result == 0:
            self._button_f5.configure(state='normal')
        # Output message
        self._view_message(result)
        
        # Append message for specific status codes
        list_message = []
        # Normal completion
        if result == 0:
            list_message.append('\r\n')
            list_message.append('　・PDF file: ' + pdf_path)
        # Abnormal completion during check
        elif result == -7107:
            # Read CSV file
            _input_path = self.settings['installdir'] + '/input/FormB_ZipFileList.csv'
            _results_data = ctrlCsv.read_csv(_input_path)
            _last_row = _results_data[-1]
            #
            list_message.append('\r\n')
            list_message.append('　・ZIP file: ' + _last_row["FilePath"] + '\r\n')
        # Append message
        self._append_message("".join(list_message))
    
    # F3 button
    def _click_f3(self):
        print('_click_f3')
    
    # F4 button
    def _click_f4(self):
        print('_click_f4')
    
    # F5 button
    def _click_f5(self):
        result = 0

        _entry_data = self._get_entry_data()
        target_path = _entry_data.folderpath

        _pathlib = Path(target_path)
        target_path = str(_pathlib.parent)
        
        result = self._open_output_path(target_path)

        # Do not output messages in normal cases when opening folder
        if result != 0:
            self._view_message(result)
    
    # F6 button
    def _click_f6(self):
        print('_click_f6')

    # F7 button
    def _click_f7(self):
        print('_click_f7')

    # F8 button
    def _click_f8(self):
        print('_click_f8')

    # F9 button
    def _click_f9(self):
        _result = 0

        target_path = self.settings_filepath
        
        # Set parent folder
        _pathlib = Path(target_path)
        target_path = str(_pathlib.parent)

        _result = self._open_output_path(target_path)

        # Do not output messages in normal cases
        if _result != 0:
            self._view_message(_result)

    # F10 button
    def _click_f10(self):
        print('_click_f10')

    # F11 button
    def _click_f11(self):
        print('_click_f11')

    # F12 button
    def _click_f12(self):
        self.master.destroy()
        sys.exit(0)

    # Method-related processing --->

    def _toggle_buttons(self, correct_password):
        _result = 0

        # Switch to check mode
        if self._button_f2["state"] == 'disabled':
            # Password authentication
            # password = simpledialog.askstring('Password Authentication', 'Enter password.', show='*')
            _password = formAuth(self.master, "Password Authentication").result
            # Cancel
            if _password is None:
                _result = 9999

            # Password authentication succeeded
            elif _password == correct_password:
                try:
                    self._change_mode(1)
                    
                except Exception as err:
                    _result = -1001
            # Password authentication failed
            else:
                tk.messagebox.showerror('Password Authentication Failed', 'Incorrect password.')
                _result = 9999

        # Switch to package mode
        elif self._button_f2["state"] == 'normal':
            try:
                self._change_mode(2)
            except Exception as err:
                _result = -1002
        
        return _result
    
    def _change_mode(self, mode):
        # Package mode
        if mode == 1:
            # Change mode label
            self._label_mode.configure(
                text=self.settings['label_mode_package'],
                fg=self.settings['fg_color_package'],
                bg=self.settings['bg_color_package']
            )
            # Enable/disable buttons
            self._button_f1.configure(state='normal')
            self._button_f2.configure(state='disabled')
            self._button_f5.configure(state='disabled')
            self._button_f9.configure(state='disabled')
        # Check mode
        elif mode == 2:
            # Change mode label
            self._label_mode.configure(
                text=self.settings['label_mode_check'],
                fg=self.settings['fg_color_check'],
                bg=self.settings['bg_color_check']
            )
            # Enable/disable buttons
            self._button_f1.configure(state='disabled')
            self._button_f2.configure(state='normal')
            self._button_f5.configure(state='disabled')
            self._button_f9.configure(state='normal')
    
    # Create array data for input values
    def _get_entry_data(self):
        entry_data = structureEntrydata.EntryData(
            folderpath = self._textbox_folderpath.get(),
            targetrange = self._combobox_targetrange.get(),
            workdate = self._dateentry_workdate.get(),
            workername = self._combobox_workername.get(),
            terminalname = self._combobox_terminalname.get()
        )
        return entry_data
    
    def _get_csv_data(self):
        _entry_data = self._get_entry_data()
        _csv_data = [
            ["OverallResult","WorkDate","WorkerName","TargetProcess","TerminalName","TargetFolder","Result"],
            ["All OK", _entry_data.workdate, _entry_data.workername, _entry_data.targetrange, _entry_data.terminalname, _entry_data.folderpath, "OK"]
        ]

        return _csv_data
    
    # Check input values
    def _validate_entry(self, entry_data: structureEntrydata.EntryData):
        _result = 0
        # Target folder - Input check
        if entry_data.folderpath == "":
            _result = -1101
        # Target folder - Check folder existence
        if _result == 0:
            if not os.path.isdir(entry_data.folderpath):
                _result = -1102

        # Work target - ComboBox
        if _result == 0:
            if entry_data.targetrange == "":
                _result = -1111
        # Work date
        if _result == 0:
            try:
                datetime.strptime(entry_data.workdate, "%m/%d/%Y")
            except ValueError:
                _result = -1121
        # Department name / Worker name
        if _result == 0:
            if entry_data.workername == "":
                _result = -1131
        # Work terminal name
        if _result == 0:
            if entry_data.terminalname == "":
                _result = -1141
        
        return _result
    
    # Output input values as CSV file for report header
    def _output_csv_headervalues(self, output_path):
        _result = 0
        _csv_data = self._get_csv_data()
        _result = ctrlCsv.output_csv(output_path, _csv_data)

        return _result
    
    # Check file count and file size (MB) in target folder
    def _check_folder_limits(self, target_path, max_size, max_files):
        _result = 0

        # Skip check if both are "0"
        if max_size == 0 and max_files == 0:
            return _result
        
        total_size, file_count  = self._get_folder_stats(target_path)

        # Check file size
        if max_size != 0 and total_size > max_size:
                _result = -1201
        if max_files != 0 and file_count > max_files:
                _result = -1202

        return _result
    
    # Get file count and file size (MB) in target folder
    def _get_folder_stats(self, folder_path):
        total_bytes = 0
        file_count = 0

        for current_dir, _, file_lists in os.walk(folder_path):
            total_bytes += sum(os.path.getsize(os.path.join(current_dir, file)) for file in file_lists)
            file_count += len(file_lists)
        
        return total_bytes, file_count

    # Packaging
    def _packaging_data(self, script_path, hash_algorithm, input_path, output_path):
        args = [
            '-Pack',
            '-Hash', hash_algorithm,
            '-InputPath', input_path,
            '-OutputPath', output_path
        ]
        result = ctrlBatch.exe_powershell(script_path, *args)

        return result, output_path
    
    def _get_unique_filename(self, file_path, max_attempts=0):
        # Return as is if file does not exist
        if not os.path.exists(file_path):
            return file_path

        base, ext = os.path.splitext(file_path)

        # No attempts (repeat indefinitely to set unique file path)
        if max_attempts == 0:
            counter = 1
            while os.path.exists(f"{base}-{counter}{ext}"):
                counter += 1

            return f"{base}-{counter}{ext}"            
        
        # With attempts
        else:
            for counter in range(1, max_attempts + 1):
                new_file_path = f"{base}-{counter}{ext}"
                if not os.path.exists(new_file_path):
                    return new_file_path
            raise Exception(f"Exceeded {max_attempts} attempts. Failed to generate a suitable filename.")
    
    def _output_csv_formalistbodyvalues(self, input_path, output_path):
        _result = 0

        _result = ctrlCsv.extract_xml_to_csv(input_path, output_path)

        return _result
    
    def _output_csv_formblist_bodyvalues(self, filepath_lists, output_path):
        _result = 0

        _result = ctrlCsv.extract_xmls_to_csv(filepath_lists, output_path)

        return _result

    # Print Form A list
    def _print_formalists(self, script_path, script_args):
        _result = 0

        _result = ctrlExcel.testexcel()

        if _result == 0:
            args = [
                '-FormA',
                '-OutputPath', script_args['output_path'],
                '-RootPath', script_args['root_path'],
                '-DataMapping_Header', script_args['datamapping_header'],
                '-DataMapping_Body', script_args['dataMapping_body'],
                '-FormA_Template', script_args['forma_template'],
                '-FormA_HeaderValues', script_args['forma_headerValues'],
                '-FormA_BodyValues', script_args['forma_bodyValues']
            ]
            _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    def _copy_formdata(self, from_path, to_path):
        result = 0

        try:
            shutil.copy(from_path, to_path)
        except:
            result = -1203

        return result, to_path
    
    def _print_formblists(self, script_path, script_args):
        _result = 0

        _result = ctrlExcel.testexcel()

        if _result == 0:    
            args = [
                '-FormB',
                '-OutputPath', script_args['output_path'],
                '-RootPath', script_args['root_path'],
                '-DataMapping_Header', script_args['datamapping_header'],
                '-DataMapping_Body', script_args['dataMapping_body'],
                '-FormB_Template', script_args['formb_template'],
                '-FormB_HeaderValues', script_args['formb_headerValues'],
                '-FormB_BodyValues', script_args['formb_bodyValues']
            ]
            _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    def _open_output_path(self, target_path):
        _result = 0

        _result = ctrlBatch.open_folder(target_path)

        return _result
    
    # Check packaged data
    def _check_data(self, script_path, input_path, output_path):
        args = [
            '-InputPath', input_path,
            '-OutputPath', output_path
        ]
        _result = ctrlBatch.exe_powershell(script_path, *args)

        return _result

    # Display processing result and message
    def _view_message(self, result):
        # Set processing result label
        if result == 9999:
            self._label_result['text'] = self.settings['label_result_init']
            self._label_result['fg'] = self.settings['fb_color_normal']
        elif result == 9000:
            self._label_result['text'] = ' Executing '
            self._label_result['fg'] = self.settings['fb_color_ok']
        elif 9999 > result >= 0:
            self._label_result['text'] = '  Completed  '
            self._label_result['fg'] = self.settings['fb_color_ok']
        else:
            self._label_result['text'] = '  Abnormal  '
            self._label_result['fg'] = self.settings['fb_color_ng']
        # Set message
        datestr = ctrlString.now_label()
        messages = ctrlMessage.get_message(result)
        messages = "{} {}".format(datestr, messages[1])
        self._text_messages.configure(state='normal')
        self._text_messages.delete('1.0', 'end')
        self._text_messages.insert('1.0', messages)
        self._text_messages.configure(state='disabled')
        # Update
        self._label_result.update()
        self._text_messages.update()
    
    def _append_message(self, messages):
        self._text_messages.configure(state='normal')
        self._text_messages.insert(tk.END, messages)
        self._text_messages.configure(state='disabled')
        # Update
        self._text_messages.update()
