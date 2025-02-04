import customtkinter as ctk
import win32com.client
import tkinter as tk


class Window():
    def __init__(self, windowTitle: str = "HardwareMonitor", windowGeometry: str = "500x600") -> None:
        self.root = ctk.CTk()
        self.windowTitle = windowTitle
        self.windowGeometry = windowGeometry
        self.setupWindow()

    def setupWindow(self, windowTitle: str = "HardwareMonitor", windowGeometry: str = "500x600") -> None:
        self.root.title(windowTitle)
        self.root.geometry(windowGeometry)


    def buildWindow(self):
        self.root.mainloop()


class MainFrame():
    def __init__(self, parent: Window) -> None:
        self.mainframe = ctk.CTkFrame(parent.root, fg_color="transparent")
        self.mainframe.grid_rowconfigure(0, weight=1)
        self.mainframe.grid_columnconfigure(0, weight=1)
        self.mainframe.pack(fill="both", expand=True)


class UIelements():
    def __init__(self, parent: Window, mainFrame: MainFrame) -> None:
        self.parent = parent
        self.mainFrame = mainFrame.mainframe
        self.textbox = ctk.CTkTextbox(self.mainFrame, activate_scrollbars=False)
        self.scrollbar = ctk.CTkScrollbar(self.mainFrame, command=self.textbox.yview)
        self.textbox.configure(font=("Calibri", 15))
        self.setupElements()

    def setupElements(self):
        self.textbox.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.textbox.configure(yscrollcommand=self.scrollbar.set)
        self.textbox.bind("<Button-3>", self.show_context_menu)
        self.textbox.bind("<Control-C>", self.copy_text)
        

    def copy_text(self):
        selected_text = self.textbox.get(tk.SEL_FIRST, tk.SEL_LAST)
        self.parent.root.clipboard_clear()
        self.parent.root.clipboard_append(selected_text)

    def show_context_menu(self, event):
        context_menu = tk.Menu(self.parent.root, tearoff=0)
        context_menu.add_command(label="Копіювати", command=self.copy_text)
        context_menu.tk_popup(event.x_root, event.y_root)


class CPU():
    def get(self, cpus) -> str:
        info = ""
        for processor in cpus:
            info = f"""Процесор < {processor.Name.strip()} >
        Кількість ядер: {str(processor.NumberOfCores)}
        Виробник: {processor.Manufacturer}
        Частота процесора: {str(processor.MaxClockSpeed)} MHz
        Сокет: {processor.SocketDesignation}"""
            
        return info


class GPU():
    def get(self, gpus) -> str:
        info = ""
        for video in gpus:

            info = f'''Відеокарта < {str(video.Name)} >
        Виробник: {str(video.VideoProcessor)}
        Версія відеодрайвера: {str(video.DriverVersion)}
        Поточна роздільна здатність: {str(video.CurrentHorizontalResolution)}x{str(video.CurrentVerticalResolution)}'''

        return info


class Motherboard():
    def get(self, motherboard) -> str:
        info = ""

        if motherboard:
            info = f"Материнська плата < {str(motherboard[0].Manufacturer)} {str(motherboard[0].Product)} >"

            if motherboard[0].Version.strip() != "":
                info += f"\n        Версія материнської плати: {str(motherboard[0].Version)}"

            if motherboard[0].SerialNumber.strip() != "":
                info += f'''\n        Серійний номер материнської плати: {str(motherboard[0].SerialNumber)}
        Чіпсет материнської плати: {str(motherboard[0].Caption)}'''
        else:
            info = "Інформація про материнську плату не знайдена."
        
        return info


class RAM():
    def __init__(self) -> None:
        self.total_ram_volume_in_gb = 0

    def get(self, ram_sticks, decor: str | int | None) -> str:
        info = "Оперативна пам'ять\n"

        for stick in ram_sticks:
            ram_stick_capacity_in_gb = int(stick.Capacity) / 1073741824
            self.total_ram_volume_in_gb += ram_stick_capacity_in_gb
            info += f'''        {decor} {str(stick.Manufacturer.strip())} {str(stick.PartNumber.strip())} ({ram_stick_capacity_in_gb}ГБ, {str(stick.Speed)}Hz)\n'''
            
        info += f"\nЗагальний обсяг оперативної пам'яті: {str(self.total_ram_volume_in_gb)}ГБ"
        return info


class DISK():
    def get(self, disks, decor) -> str:
        info = "Диски\n"

        for disk in disks:
                info += f'''        {decor} {str(disk.Model)}: {str(int(int(disk.Size) / (1024**3)))}ГБ\n'''

        return info
    

class HardwareMonitor:
    def __init__(self, update_ui_callback=None) -> None:
        self.separator = "\n- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\n"
        self.textDecorator = "‣"
        self.update_ui_callback = update_ui_callback

    def get_hardwareInfo(self) -> None:
        try:
            import pythoncom
            pythoncom.CoInitialize ()
            self.wmi = win32com.client.GetObject("winmgmts:")
        except Exception as e:
            self.update_ui_callback(f"Error accessing WMI: {e}")
        
        hardwareInfo: str = (
                        CPU().get(self.wmi.ExecQuery("SELECT * FROM Win32_Processor")) + 
                        self.separator +
                        GPU().get(self.wmi.ExecQuery("SELECT * FROM Win32_VideoController")) +
                        self.separator +
                        Motherboard().get(self.wmi.ExecQuery("SELECT * FROM Win32_BaseBoard")) +
                        self.separator +
                        RAM().get(self.wmi.ExecQuery("SELECT * FROM Win32_PhysicalMemory"), self.textDecorator) +
                        self.separator +
                        DISK().get(self.wmi.ExecQuery("SELECT * FROM Win32_DiskDrive"), self.textDecorator))

        self.update_ui_callback(hardwareInfo)



class Application():
    def __init__(self) -> None:
        self.window = Window()
        mainFrame = MainFrame(self.window)
        self.UiElements = UIelements(self.window, mainFrame)
        self.hardwareMonitor = HardwareMonitor(self.update_text)

    def setText(self, text: str ="We're getting info..."):
        self.UiElements.textbox.insert("1.0", text)

    def update_text(self, text: str):
        self.UiElements.textbox.delete("1.0", tk.END)
        self.UiElements.textbox.insert("1.0", text)
    
    def run(self):
        self.setText()

        import threading
        threading.Thread(target=self.hardwareMonitor.get_hardwareInfo, daemon=True).start()

        self.window.buildWindow()





if __name__ == "__main__":
    app = Application()
    app.run()
