import tkinter as tk

class formBase(object):
    def __init__(self, master):
        # Main window
        self.master = master
        self.settings = formBase.base_settings()
        self.master.title(self.settings.get('title'))
        self.master.geometry(self.settings.get('window_size'))
        self.master.configure(bg=self.settings.get('bg_color'))

        # Area of header
        self.frame_header = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_header.pack(expand=1, fill="both", side="top")

        # Area of main
        self.frame_body = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_body.pack(expand=1, fill="both", side="top")

        # Area of button
        self.frame_footer = tk.Frame(
            self.master,
            bg=self.settings['bg_color'],
            padx=12,
            pady=12
        )
        self.frame_footer.pack(expand=1, fill="both", side="bottom")

    def base_settings():
        settings = {}
        settings["title"] = ''
        settings["window_size"] = '960x540'
        settings["font_label_h1"] = ['Segoe UI Mono', 15, 'bold']
        settings["font_label_h2"] = ['Segoe UI Mono', 14, 'bold']
        settings["font_label_body"] = ['Segoe UI Mono', 12]
        settings['font_button'] = ['Segoe UI Mono', 12]
        settings['font_text'] = ['Segoe UI Mono', 15]
        settings['font_listbox'] = ['Segoe UI Mono', 12]
        settings['font_combobox'] = ['Segoe UI Mono', 12]
        settings['font_cal'] = ['Segoe UI Mono', 12]
        #   ウィンドウ設定
        settings["bg_color"] = '#f5f5f5'

        return settings

    def start(self):
        self.master.resizable(False, False)
        self.master.mainloop()

    def exit(self):
        self.master.exit()
