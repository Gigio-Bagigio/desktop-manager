import os
import customtkinter as ctk
from tkinter import messagebox, filedialog
import winreg
import shutil as sht
import subprocess
from hPyT import opacity
from PIL import Image
import time
import threading
import pystray  
import sys
import tkinter as tk

user = os.getlogin()
basePath = f"C:/Users/{user}/DesktopProfiles"
profileFile = f"{basePath}/active_profile.txt"
tempMarkerFile = f"{basePath}/.temp_active"
defaultDesktop = f"C:/Users/{user}/Desktop"

resourcesPath = f"{basePath}/resources"
os.makedirs(resourcesPath, exist_ok=True)

# colors
PRIMARY_COLOR = "#1f538d"
ACCENT_COLOR = "#3a7ebf"
DANGER_COLOR = "#e63946"
SUCCESS_COLOR = "#2a9d8f"
TEMP_COLOR = "#f4a261"

def apply_theme(theme_path):
    if not os.path.exists(theme_path):
        print(f"Theme '{theme_path}' doesn't exist.")
        return False
    try:
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "CurrentTheme", 0, winreg.REG_SZ, theme_path)
            
        command = f'''$signature = @"
[DllImport("uxtheme.dll", CharSet = CharSet.Auto)]
public static extern int SetSystemVisualStyle(string pszFilename, string pszColor, string pszSize, int dwReserved);
"@

$uxtheme = Add-Type -MemberDefinition $signature -Name SysTheme -Namespace Win32Functions -PassThru
$uxtheme::SetSystemVisualStyle("{theme_path}", $null, $null, 0)'''
        
        result = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)
        if result.returncode == 0:
            print(f"Theme '{theme_path}' applied succesfully.")
            return True
        else:
            print(f"Error applying theme: {result.stderr}")
            
            command = f'start-process -filepath "{theme_path}"; timeout /t 02; taskkill /im systemsettings.exe /f'
            subprocess.run(["powershell", "-Command", command], check=True)
            return True
            
    except Exception as e:
        print(f"Error applying theme: {e}")

        try:
            command = f'start-process -filepath "{theme_path}"; timeout /t 02; taskkill /im systemsettings.exe /f'
            subprocess.run(["powershell", "-Command", command], check=True)
            return True
        except Exception as e2:
            print(f"Fallback also failed: {e2}")
            return False

def get_current_theme_path():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes", 0, winreg.KEY_READ)
        current_theme, _ = winreg.QueryValueEx(key, "CurrentTheme")
        winreg.CloseKey(key)
        return current_theme
    except Exception as e:
        print("Error reading log: ", e)
        return None

os.makedirs(basePath, exist_ok=True)

# Profile management functions

def get_profiles():
    return [f.replace(".reg", "") for f in os.listdir(basePath) if f.endswith(".reg")]

def get_active_profile():
    if os.path.exists(tempMarkerFile):
        return "temp"
        
    if os.path.exists(profileFile):
        with open(profileFile, "r") as f:
            return f.readline().strip()
    return "default" # if doesn't exist return default

def set_active_profile(profile):
    # if the profile is "temp", create a marker file
    if profile == "temp":
        with open(tempMarkerFile, "w") as f:
            f.write("temporary desktop active")

    elif os.path.exists(tempMarkerFile):
        # if we are switching from temp to another profile, remove the marker
        os.remove(tempMarkerFile)
    
    # save profile
    with open(profileFile, "w") as f:
        f.write(profile)

def create_toast_notification(message, duration=3):
    toast = ctk.CTkToplevel()
    toast.attributes("-topmost", True)
    toast.overrideredirect(True)
    toast.configure(fg_color=ACCENT_COLOR)

    width = 300
    height = 50
    screen_width = toast.winfo_screenwidth()
    screen_height = toast.winfo_screenheight()
    x = screen_width - width - 20
    y = screen_height - height - 60
    
    toast.geometry(f"{width}x{height}+{x}+{y}")
    
    ctk.CTkLabel(toast, text=message, font=ctk.CTkFont(size=14), text_color="white").pack(expand=True)
    
    def close_toast():
        time.sleep(duration)
        toast.destroy()
    
    threading.Thread(target=close_toast, daemon=True).start()
    return toast

def save_layout(profile=""): 

    if profile == "temp": # we don't need to save the layout for the temp profile
        return
        
    ####### theme #######
    try:
        activeThemePath = get_current_theme_path()
        if activeThemePath is None:
            activeThemePath = r"C:\WINDOWS\resources\Themes\themeD.theme" # default theme
        sht.copy(activeThemePath, fr"{basePath}/{profile}.theme")

        lines = []
        
        profileConfigurationFilePath = ""

        if profile != "default":
            profileConfigurationFilePath = f"{basePath}/{profile}_desktop.txt"
            if os.path.exists(profileConfigurationFilePath):
                with open(profileConfigurationFilePath, "r") as f:
                    lines = f.readlines() 
        else:
            profileConfigurationFilePath = f"{basePath}/desktop.txt"
            if not os.path.exists(profileConfigurationFilePath):
                with open(profileConfigurationFilePath, "w") as f:
                    f.write(defaultDesktop)
            with open(profileConfigurationFilePath, "r") as f:
                lines = f.readlines()

        with open(profileConfigurationFilePath, "w") as f:
            f.writelines(lines)
        print(lines)    
        print(profileConfigurationFilePath)

    except Exception as e:
        print(f"Error saving theme: {e}")

    ######################## 
    # get the icons layout file from the registry
    if profile == "":
        profile = profile_var.get()
    file_path = f'{basePath}/{profile}.reg'
    print("Saving the registry...")
    command = f'reg export "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\Shell\\Bags\\1\\Desktop" "{file_path}" /y'
    os.system(command)
    create_toast_notification(f"Layout saved for {profile}")

def update_desktop_registry(new_desktop):
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, "Desktop", 0, winreg.REG_EXPAND_SZ, new_desktop.replace('/', "\\"))
            print("Registry updated successfully.")
            return True
    except Exception as e:
        print(f"Error updating registry: {e}")
        return False


def restore_layout():
    """ restore the layout of the seleted profile """
    profile = profile_var.get()
    file_path = f'{basePath}/{profile}.reg'
    desktop_path_file = f"{basePath}/{profile}_desktop.txt"
    print(desktop_path_file)
    
    if os.path.exists(file_path):
        print("Registry import...")
        os.system(f'reg import "{file_path}"')
        
        if os.path.exists(desktop_path_file):
            with open(desktop_path_file, "r") as f:
                new_desktop = f.readline().strip()
        else:
            new_desktop = defaultDesktop   
            
        if update_desktop_registry(new_desktop):
            restart_explorer()
            create_toast_notification(f"Arrangement restored for {profile}")
        else:
            messagebox.showwarning("Error", "Unable to update registry.")
    else:
        messagebox.showwarning("Error", f"No arrangements saved for {profile}")

def restart_explorer():
    if 'status_label' in globals():
        status_label.configure(text="Explorer restarting in progress...")
        root.update()
    
    os.system("taskkill /f /im explorer.exe")
    time.sleep(1)
    os.system("start explorer.exe")
    
    if 'status_label' in globals():
        status_label.configure(text="Explorer restarted successfully")
        root.after(3000, lambda: status_label.configure(text=""))

def create_profile():
    nuovo_profilo = profile_entry.get().strip() # new profile name
    if not nuovo_profilo:
        messagebox.showwarning("Error", "Enter a name for the profile")
        return
        
    if nuovo_profilo == "temp":
        messagebox.showwarning("Reserved name", "'temp' is a name reserved for the temporary profile. Choose another name.")
        return
        
    if nuovo_profilo in get_profiles(): # check if the name is already used
        result = messagebox.askyesno("Existing profile", f"Profile '{nuovo_profilo}' already exists. Do you want to overwrite it?")
        if not result:
            return
    
    desktop_path = filedialog.askdirectory(title="Select the Desktop folder for this profile") or defaultDesktop

    with open(f"{basePath}/{nuovo_profilo}.reg", "w") as f:
        f.write("")  # Create a empty file
    
    with open(f"{basePath}/{nuovo_profilo}_desktop.txt", "w") as f:
        f.write(desktop_path)  # save the new desktop path
        
    # 
    currentTheme = get_current_theme_path()
    if currentTheme is None:
        currentTheme = r"C:\WINDOWS\resources\Themes\themeD.theme"
    sht.copy(currentTheme, fr"{basePath}/{nuovo_profilo}.theme")

    update_profiles()
    profile_var.set(nuovo_profilo)
    change_profile(nuovo_profilo)
    create_toast_notification(f"Profile '{nuovo_profilo}' created succesfly!")
    profile_entry.delete(0, 'end') 
    updateTrayIcon()

def create_profile_ui(duration=30):
    create_profile_w = ctk.CTkToplevel()
    create_profile_w.attributes("-topmost", True)
    create_profile_w.overrideredirect(True)
    create_profile_w.configure(fg_color=ACCENT_COLOR)

    width = 300
    height = 100
    screen_width = create_profile_w.winfo_screenwidth()
    screen_height = create_profile_w.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    create_profile_w.geometry(f"{width}x{height}+{x}+{y}")

    # Canvas for draw the X on the circle
    canvas = tk.Canvas(create_profile_w, width=24, height=24, highlightthickness=0, bg=ACCENT_COLOR, bd=0)
    canvas.place(x=width - 30, y=6)

    # Draw a circle with white border
    canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5, fill=ACCENT_COLOR)

    # white X
    line_width = 1.5
    canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
    canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)

    def close_create_profile_w(event=None):
        create_profile_w.destroy()

    canvas.bind("<Button-1>", close_create_profile_w)

    # hover effect for the close button
    def on_enter(event):
        canvas.delete("all")
        canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5)#, fill="#ff3333")
        canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
        canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)
    
    def on_leave(event):
        canvas.delete("all")
        canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5, fill=ACCENT_COLOR)
        canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
        canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)

    # Bind per hover
    canvas.bind("<Enter>", on_enter)
    canvas.bind("<Leave>", on_leave)

    ctk.CTkLabel(create_profile_w, text="create profile", font=ctk.CTkFont(size=14), text_color="white").pack(expand=True)
    profileName = tk.Entry(create_profile_w, bg="#6aa2d3", border=0)
    profileName.pack(expand=True)
    ctk.CTkButton(
        create_profile_w,
        text="create profile",
        border_color="white",
        border_width=2,
        fg_color=ACCENT_COLOR,
        command=lambda: create_profile_from_tray(profileName.get().strip(), create_profile_w)
    ).pack(expand=True)
    
    # automatic close
    def auto_close():
        time.sleep(duration)
        if create_profile_w.winfo_exists():
            create_profile_w.destroy()
    
    threading.Thread(target=auto_close, daemon=True).start()
    create_profile_w.mainloop()

def create_profile_from_tray(profileName, ui):
    ui.destroy()
    new_profile = profileName
    if not new_profile:
        messagebox.showwarning("Error", "Insert a name for the profile")
        return
        
    if new_profile == "temp":
        messagebox.showwarning("Reserved name", "'temp' is a name reserved for the temporary profile. Choose another name.")
        return
        
    if new_profile in get_profiles():
        result = messagebox.askyesno("Existing Profile", f"The profile '{new_profile}' already exists. Do you want to overwrite it?")
        if not result:
            return
    
    desktop_path = filedialog.askdirectory(title="Select the Desktop folder for this profile") or defaultDesktop

    with open(f"{basePath}/{new_profile}.reg", "w") as f:
        f.write("")  # Create an empty file
    
    with open(f"{basePath}/{new_profile}_desktop.txt", "w") as f:
        f.write(desktop_path)  # Save the new desktop path
        
    currentTheme = get_current_theme_path()
    if currentTheme is None:
        currentTheme = r"C:\WINDOWS\resources\Themes\themeD.theme"
    sht.copy(currentTheme, fr"{basePath}/{new_profile}.theme")

    update_profiles()
    change_profile(new_profile)
    create_toast_notification(f"Profile '{new_profile}' successfully created!")
    updateTrayIcon()

def updateTrayIcon():
    try:
        tray_icon.menu = create_tray_menu()
    except Exception as e:
        print(f"Error updating tray icon: {e}")

def load_profile():

    active_profile = get_active_profile()
    profile = profile_var.get() # get the selected profile from the combobox
    
    # if the profile is "temp", we don't need to save the layout
    if active_profile != "temp" and active_profile != profile:
        save_layout(active_profile)
    
    if active_profile == "temp":
        clean_temp_profile()
    
    set_active_profile(profile)
    restore_layout()

    themePath = fr"{basePath}/{profile}.theme"
    if not os.path.exists(themePath):
        themePath = get_current_theme_path()
    if themePath and os.path.exists(themePath):
        apply_theme(themePath)
    
    status_label.configure(text=f"Profile '{profile}' loaded correctly!")
    root.after(3000, lambda: status_label.configure(text=""))
    updateTrayIcon()

def delete_profile():

    profile_to_delete = profile_var.get()
    if profile_to_delete == "default":
        messagebox.showwarning("Error", "You can't delete the default profile!")
        return
        
    if profile_to_delete == "temp":
        messagebox.showwarning("Error", "You can't delete the temporary profile! It will be deleted when you load another profile.")
        return
    
    if profile_to_delete == get_active_profile():
        result = messagebox.askyesno("Attention", "You are about to delete the active profile. Continue?")
        if not result:
            return
    
    result = messagebox.askyesno("Confirm", f"Are you sure you want to delete the profile '{profile_to_delete}'?")
    if not result:
        return

    try:
        os.remove(f"{basePath}/{profile_to_delete}.reg")
        os.remove(f"{basePath}/{profile_to_delete}_desktop.txt")
        if os.path.exists(f"{basePath}/{profile_to_delete}.theme"):
            os.remove(f"{basePath}/{profile_to_delete}.theme")
    except FileNotFoundError as e:
        print(f"File not found during deletion: {e}")
    
    # Se il profilo eliminato era quello attivo, passa a default
    if profile_to_delete == get_active_profile():
        set_active_profile("default")
    
    update_profiles()
    create_toast_notification(f"Profile '{profile_to_delete}' deleted!")
    updateTrayIcon()

def rename_profile(profile, newName, ui):
    ui.destroy()
    try:
        os.rename(f"{basePath}/{profile}.reg", f"{basePath}/{newName.replace(" ", "_")}.reg")
        os.rename(f"{basePath}/{profile}_desktop.txt", f"{basePath}/{newName.replace(" ", "_")}_desktop.txt")
        os.rename(f"{basePath}/{profile}.theme", f"{basePath}/{newName.replace(" ", "_")}.theme")
    except Exception as e:
        print(e)
    updateTrayIcon()

    import json

    with open("profiles.json", "r") as f:
        data = json.load(f)

    
    if profile in data:
        data[newName] = data[profile]
        del data[profile]      
    else:
        print(f"Profile '{profile}' not found.")

    # Save the JSON
    with open("profiles.json", "w") as f:
        json.dump(data, f, indent=2)

def rename_profile_ui(duration=30, profile=""):
    if profile == "":
        try:
            profile = profile_var.get()
            print(profile)
        except:
            return
    create_profile_w = ctk.CTkToplevel()
    create_profile_w.attributes("-topmost", True)
    create_profile_w.overrideredirect(True)
    create_profile_w.configure(fg_color=ACCENT_COLOR)

    width = 300
    height = 100
    screen_width = create_profile_w.winfo_screenwidth()
    screen_height = create_profile_w.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    create_profile_w.geometry(f"{width}x{height}+{x}+{y}")

    canvas = tk.Canvas(create_profile_w, width=24, height=24, highlightthickness=0, bg=ACCENT_COLOR, bd=0)
    canvas.place(x=width - 30, y=6)

    canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5, fill=ACCENT_COLOR)

    line_width = 1.5
    canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
    canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)

    def close_create_profile_w(event=None):
        create_profile_w.destroy()

    canvas.bind("<Button-1>", close_create_profile_w)

    def on_enter(event):
        canvas.delete("all")
        canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5)#, fill="#ff3333")
        canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
        canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)
    
    def on_leave(event):
        canvas.delete("all")
        canvas.create_oval(2, 2, 22, 22, outline="white", width=1.5, fill=ACCENT_COLOR)
        canvas.create_line(8, 8, 16, 16, fill="white", width=line_width, capstyle=tk.ROUND)
        canvas.create_line(16, 8, 8, 16, fill="white", width=line_width, capstyle=tk.ROUND)

    canvas.bind("<Enter>", on_enter)
    canvas.bind("<Leave>", on_leave)

    ctk.CTkLabel(create_profile_w, text=f"rename profile {profile}", font=ctk.CTkFont(size=14), text_color="white").pack(expand=True)
    profileName = tk.Entry(create_profile_w, bg="#6aa2d3", border=0)
    profileName.pack(expand=True)
    ctk.CTkButton(
        create_profile_w,
        text="rename profile",
        border_color="white",
        border_width=2,
        fg_color=ACCENT_COLOR,
        command=lambda: rename_profile(profile, profileName.get().strip(), create_profile_w)
    ).pack(expand=True)
    
    def auto_close():
        time.sleep(duration)
        if create_profile_w.winfo_exists():
            create_profile_w.destroy()
    
    threading.Thread(target=auto_close, daemon=True).start()
    create_profile_w.mainloop()

def delete_specific_profile(profile):
    profile_to_delete = profile
    if profile_to_delete == "default":
        messagebox.showwarning("Error", "You can't delete the default profile!")
        return
        
    if profile_to_delete == "temp":
        messagebox.showwarning("Error", "You can't delete the temporary profile!")
        return
    
    if profile_to_delete == get_active_profile():
        result = messagebox.askyesno("Attention", "You are about to delete the active profile. Continue?")
        if not result:
            return
    
    result = messagebox.askyesno("Confirm", f"Are you sure you want to delete the profile '{profile_to_delete}'?")
    if not result:
        return

    try:
        os.remove(f"{basePath}/{profile_to_delete}.reg")
        os.remove(f"{basePath}/{profile_to_delete}_desktop.txt")
        if os.path.exists(f"{basePath}/{profile_to_delete}.theme"):
            os.remove(f"{basePath}/{profile_to_delete}.theme")
    except FileNotFoundError as e:
        print(f"File not found during deletion: {e}")
    
    # if the deleted profile was the active one, switch to default profile
    if profile_to_delete == get_active_profile():
        set_active_profile("default")
    
    update_profiles()
    create_toast_notification(f"Profile '{profile_to_delete}' deleted!")
    updateTrayIcon()


def update_profiles(): # update the list of profiles in the UI
    profili = get_profiles()
    if "default" not in profili:
        profili.append("default")

    if "default" in profili:
        profili.remove("default")
        profili = ["default"] + profili
    
    active = get_active_profile()
    if active == "temp" and "temp" not in profili:
        profili = ["temp"] + profili
    elif "temp" in profili and active != "temp":
        profili.remove("temp")

    # update the combobox with the profiles
    for widget in profiles_frame.winfo_children():
        widget.destroy()
    
    for profile in profili:
        is_active = profile == active
        
        # special color for the temporary profile
        if profile == "temp":
            fg_color = TEMP_COLOR if is_active else "transparent"
            hover_color = "#e76f51"
        else:
            fg_color = PRIMARY_COLOR if is_active else "transparent"
            hover_color = ACCENT_COLOR
        
        profile_btn = ctk.CTkButton(
            profiles_frame, 
            text=profile, 
            fg_color=fg_color,
            hover_color=hover_color,
            command=lambda p=profile: change_profile(p)
        )
        profile_btn.pack(side="left", padx=5, pady=5)
    
    # update the label
    profile_var.set(active)
    change_profile(active)

def change_profile(profile): # update the interface when a profile is selected
    profile_var.set(profile)
    
    # update the label with the selected profile
    profile_label.configure(text=f"Profile: {profile}")
    
    # load the profile information
    try:
        desktop_path = defaultDesktop
        if profile != "default":
            desktop_path_file = f"{basePath}/{profile}_desktop.txt"
            if os.path.exists(desktop_path_file):
                with open(desktop_path_file, "r") as f:
                    desktop_path = f.readline().strip()
        
        info_label.configure(text=f"Desktop Folder: {desktop_path}")
        
        # disable the delete button for default and temp profiles
        if profile in ["default", "temp"]:
            delete_button.configure(state="disabled")
        else:
            delete_button.configure(state="normal")
        
        if profile == "temp":
            temp_label.configure(text="Active temporary profile")
            temp_label.pack(before=info_label, pady=5)
        else:
            temp_label.pack_forget()
            
    except Exception as e:
        print(f"Error loading profile information: {e}")
        info_label.configure(text="Unable to load profile information")
    updateTrayIcon()

def open_temporary_desktop(folder_path=None):
    # save the current layout
    active_profile = get_active_profile()
    if active_profile != "temp":
        save_layout(active_profile)
    
    if folder_path is None:
        desktop_path = filedialog.askdirectory(title="Select the folder to use as the temporary Desktop")
        if not desktop_path:
            return  # the user canceled the dialog
    else:
        desktop_path = folder_path
    
    # save the temporary desktop path
    with open(f"{basePath}/temp_desktop.txt", "w") as f:
        f.write(desktop_path)  # save the new desktop path
    
    set_active_profile("temp")
    
    # Update the registry
    if update_desktop_registry(desktop_path):
        restart_explorer()
        create_toast_notification(f"Temporary desktop activated: {os.path.basename(desktop_path)}")
        
        # update the interface 
        if 'root' in globals():
            update_profiles()
            change_profile("temp")
    else:
        if 'messagebox' in globals():
            messagebox.showerror("Error", "Unable to update system registry")
        else:
            print("Error: Unable to update system registry")

def reset_desktop_to_default():
    # save the current layout
    active_profile = get_active_profile()
    if active_profile != "temp":
        save_layout(active_profile)
    
    set_active_profile("default")
    
    # Update the registry
    if update_desktop_registry(defaultDesktop):
        restart_explorer()
        create_toast_notification("Default desktop restored")
        
        # update the interface
        update_profiles()
        change_profile("default")
    else:
        messagebox.showerror("Error", "Unable to update system registry")

def check_temp_profile(): # check if a temporary profile is active at startup
    if os.path.exists(tempMarkerFile):
        result = messagebox.askyesno(
            "Temporary profile detected", 
            "A temporary profile active since the last session was detected.\n"
            "Do you want to restore the default profile?"
        )
        if result:
            os.remove(tempMarkerFile)
            set_active_profile("default")
            update_desktop_registry(defaultDesktop)
            restart_explorer()
            return "default"
    return get_active_profile()

# functions for the tray icon
def show_window():
    root.after(0, root.deiconify)
    root.lift()

def hide_window():
    root.withdraw()
    
def exit_app():
    if tray_icon:
        tray_icon.stop()
    root.destroy()
    sys.exit()
    

def create_tray_menu():
    profiles = get_profiles()
    active_profile = get_active_profile()

    menu_items = []

    menu_items.append(pystray.MenuItem(f"Active profile: {active_profile}", None, enabled=False))
    menu_items.append(pystray.MenuItem("Open Manager", show_window))
    menu_items.append(pystray.Menu.SEPARATOR)
    menu_items.append(pystray.MenuItem("Create Profile", create_profile_ui))
    # "change profile"  button with the list of profiles
    profile_items = []
    for profile in profiles:
        # function to switch profile
        def create_profile_switcher(p):
            def switch_profile():
                current = get_active_profile()
                if current != "temp":
                    save_layout(current)
                
                set_active_profile(p)
                profile_var.set(p)
                
                file_path = f'{basePath}/{p}.reg'
                desktop_path_file = f"{basePath}/{p}_desktop.txt"
                
                if os.path.exists(file_path):
                    os.system(f'reg import "{file_path}"')
                    
                    if os.path.exists(desktop_path_file):
                        with open(desktop_path_file, "r") as f:
                            new_desktop = f.readline().strip()
                    else:
                        new_desktop = defaultDesktop
                        
                    if update_desktop_registry(new_desktop):
                        restart_explorer()
                        create_toast_notification(f"Profile '{p}' loaded")
                        
                        # reload the tray icon menu for show the changes
                        tray_icon.menu = create_tray_menu()
                                        
                    themePath = fr"{basePath}/{p}.theme"
                    if not os.path.exists(themePath):
                        themePath = get_current_theme_path()
                    if themePath and os.path.exists(themePath):
                        apply_theme(themePath)

                    
            return switch_profile
        
        def eliminaProfile(profile):
            def deleteProfile():
                delete_specific_profile(profile)
            return deleteProfile
        def rinominaProfile(profile):
            def rinominaProfilo():
                rename_profile_ui(profile=profile)
            return rinominaProfilo    
        
        def goToDir(profile):
            def openFolder():
                desktop_path_file = f"{basePath}/{profile}_desktop.txt"
                if os.path.exists(desktop_path_file):
                    with open(desktop_path_file, "r") as f:
                        new_desktop = f.readline().strip()
                        os.startfile(new_desktop)
            return openFolder
        
        optionProfile = [pystray.MenuItem("load", create_profile_switcher(profile)), 
                         pystray.MenuItem("delete", eliminaProfile(profile)), 
                         pystray.MenuItem("rename", rinominaProfile(profile)), 
                         pystray.MenuItem("open profile folder",goToDir(profile))]

        profile_items.append(pystray.MenuItem(
            profile, 
            pystray.Menu(*optionProfile),#create_profile_switcher(profile),
            checked=lambda item, p=profile: get_active_profile() == p,
            radio=True
        ))
    
    menu_items.append(pystray.MenuItem("Profiles", pystray.Menu(*profile_items)))
    menu_items.append(pystray.Menu.SEPARATOR)
    menu_items.append(pystray.MenuItem("Temporary Desktop", lambda: (show_window(), open_temporary_desktop())))
    #menu_items.append(pystray.MenuItem("Restore Default Desktop", lambda: (reset_desktop_to_default(), tray_icon.menu = create_tray_menu()))) 
    menu_items.append(pystray.Menu.SEPARATOR)
    menu_items.append(pystray.MenuItem("exit", exit_app))
    
    return pystray.Menu(*menu_items)

def setup_tray_icon():
    icon_image = Image.open('icon.ico')
    # create the tray icon
    icon = pystray.Icon(
        "DesktopProfileManager",
        icon_image,
        "Desktop Profile Manager",
        menu=create_tray_menu()
    )
    return icon

def minimize_to_tray():
    root.withdraw()  # hide the window
    create_toast_notification("Desktop Profile Manager is still running in tray", duration=5)

def clean_temp_profile(): # clean up the files of the temp profile
    temp_files = [
        f"{basePath}/temp.reg",
        f"{basePath}/temp_desktop.txt",
        f"{basePath}/temp.theme",
        tempMarkerFile
    ]
    
    for file_path in temp_files:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"Removed temporary file: {file_path}")
            except Exception as e:
                print(f"Error removing file {file_path}: {e}")

def register_context_menu(): 
    # with this if you click with the right click 
    # of the mouse on a folder will appear the option 
    # to create a profile (/ temp profile) from that folder 
    try:
        # get the executable path
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
            exe_path = os.path.abspath(exe_path)
            print(f"Executable found: {exe_path}")
            
        else: # if it isn't compiled we need the script path
            # Python Script
            exe_path = sys.argv[0]
            exe_path = f'"{sys.executable}" "{os.path.abspath(exe_path)}"'
            print(f"Script found: {exe_path}")

        # option for open a folder as a temp profile
        temp_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_TempDesktop")
        winreg.SetValueEx(temp_key, "", 0, winreg.REG_SZ, "Open as Temporary Desktop")
        winreg.SetValueEx(temp_key, "Icon", 0, winreg.REG_SZ, exe_path)
        temp_cmd_key = winreg.CreateKey(temp_key, "command")
        winreg.SetValueEx(temp_cmd_key, "", 0, winreg.REG_SZ, f'{exe_path} --temp "%1"')
        
        # option for create a new profile 
        new_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_NewProfile")
        winreg.SetValueEx(new_key, "", 0, winreg.REG_SZ, "Create desktop profile")
        winreg.SetValueEx(new_key, "Icon", 0, winreg.REG_SZ, exe_path)
        new_cmd_key = winreg.CreateKey(new_key, "command")
        winreg.SetValueEx(new_cmd_key, "", 0, winreg.REG_SZ, f'{exe_path} --new "%1"')
        
        print("Options successfully registered.")
        return True
    except Exception as e:
        print(f"Error registering context menu: {e}")
        return False

def unregister_context_menu(): # remove the right click options
    try:
        # remove the keys from the registry
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_TempDesktop\command")
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_TempDesktop")
        except Exception as e:
            print(f"Advice: {e}")
            
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_NewProfile\command")
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"Directory\shell\DPM_NewProfile")
        except Exception as e:
            print(f"Advice: {e}")
        
        print("Options removed successfully.")
        return True
    except Exception as e:
        print(f"Error removing context menu: {e}")
        return False

def create_new_profile_from_folder(folder_path):
    """Crea un nuovo profilo utilizzando la cartella specificata come desktop."""
    # get the folder name and set it as profile name
    folder_name = os.path.basename(folder_path)
    profile_name = f"{folder_name}_profile"
    
    # create the profile
    with open(f"{basePath}/{profile_name}.reg", "w") as f:
        f.write("")
    
    with open(f"{basePath}/{profile_name}_desktop.txt", "w") as f:
        f.write(folder_path)  # save the desktop path
        
    currentTheme = get_current_theme_path()
    if currentTheme is None:
        currentTheme = r"C:\WINDOWS\resources\Themes\themeD.theme"
    sht.copy(currentTheme, fr"{basePath}/{profile_name}.theme")
    
    # save active layout
    active_profile = get_active_profile()
    if active_profile != "temp":
        save_layout(active_profile)
    else:
        clean_temp_profile()
    
    # set the new profile
    set_active_profile(profile_name)
    
    # update the registry
    if update_desktop_registry(folder_path):
        restart_explorer()
        create_toast_notification(f"New profile '{profile_name}' created and activated!")
    else:
        print("Error: Unable to update registry")

def process_command_line():
    if len(sys.argv) > 1:
        if sys.argv[1] == "--temp" and len(sys.argv) > 2:
            folder_path = sys.argv[2]
            open_temporary_desktop(folder_path)
            sys.exit(0)
            
        elif sys.argv[1] == "--new" and len(sys.argv) > 2:
            folder_path = sys.argv[2]
            create_new_profile_from_folder(folder_path)
            sys.exit(0)
            
        elif sys.argv[1] == "--register":
            if register_context_menu():
                print("Context menu successfully registered.")
            else:
                print("Error registering context menu.")
            sys.exit(0)
            
        elif sys.argv[1] == "--unregister":
            if unregister_context_menu():
                print("Context menu removed successfully.")
            else:
                print("Error removing context menu.")
                sys.exit(0)

# check if there are commands at the start
process_command_line()

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Desktop Profile Manager")
root.geometry("700x500")
root.minsize(600, 400)
try:
    root.iconbitmap("./icon.ico")
except:
    print("Icon file not found")

root.protocol("WM_DELETE_WINDOW", minimize_to_tray)

opacity.set(root, 0.95)

profile_var = ctk.StringVar()

# Principal Layout 
main_frame = ctk.CTkFrame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# Header
header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
header_frame.pack(fill="x", pady=(0, 10))

title_label = ctk.CTkLabel(header_frame, text="Desktop Profile Manager", font=ctk.CTkFont(size=24, weight="bold"))
title_label.pack(side="left", padx=10)

"""# Aggiungi pulsante per minimizzare nel tray
tray_button = ctk.CTkButton(
    header_frame,
    text="Minimizza nel Tray",
    command=minimize_to_tray,
    fg_color=ACCENT_COLOR,
    hover_color=PRIMARY_COLOR,
    width=150
)
tray_button.pack(side="right", padx=10)"""

# Available Profiles
profiles_label = ctk.CTkLabel(main_frame, text="Available profiles:", font=ctk.CTkFont(size=16))
profiles_label.pack(anchor="w", padx=10, pady=(5, 0))

profiles_container = ctk.CTkFrame(main_frame)
profiles_container.pack(fill="x", padx=10, pady=5)

profiles_frame = ctk.CTkFrame(profiles_container, fg_color=("#e0e0e0", "#333333"))
profiles_frame.pack(fill="x", padx=5, pady=5)

# Frame for the profile's info
content_frame = ctk.CTkFrame(main_frame)
content_frame.pack(fill="both", expand=True, padx=10, pady=10)

profile_label = ctk.CTkLabel(content_frame, text="", font=ctk.CTkFont(size=20, weight="bold"))
profile_label.pack(pady=(10, 5))

temp_label = ctk.CTkLabel(content_frame, text="Active temporary profile", 
                         font=ctk.CTkFont(size=14), 
                         text_color=TEMP_COLOR)

info_label = ctk.CTkLabel(content_frame, text="", font=ctk.CTkFont(size=14))
info_label.pack(pady=10)

# Frame for the buttons
action_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
action_frame.pack(pady=10, fill="x")

load_button = ctk.CTkButton(
    action_frame, 
    text="Upload Profile", 
    command=load_profile,
    fg_color=PRIMARY_COLOR,
    hover_color=ACCENT_COLOR,
    font=ctk.CTkFont(size=14, weight="bold"),
    width=150,
    height=40
)
load_button.pack(side="left", padx=5)

delete_button = ctk.CTkButton(
    action_frame, 
    text="Delete Profile", 
    command=delete_profile,
    fg_color=DANGER_COLOR,
    hover_color="#c1121f",
    font=ctk.CTkFont(size=14),
    width=150,
    height=40
)
delete_button.pack(side="left", padx=5)
rinomina_button = ctk.CTkButton(
    action_frame, 
    text="Rename Profile", 
    command=rename_profile_ui,
    fg_color="#b6a12b",
    hover_color="#c1121f",
    font=ctk.CTkFont(size=14),
    width=150,
    height=40
)
rinomina_button.pack(side="left", padx=5)

# temp
temp_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
temp_frame.pack(pady=10, fill="x")

temp_button = ctk.CTkButton(
    temp_frame, 
    text="Open Temporary Desktop", 
    command=open_temporary_desktop,
    fg_color=TEMP_COLOR,
    hover_color="#e76f51",
    font=ctk.CTkFont(size=14),
    width=200,
    height=40
)
temp_button.pack(side="left", padx=5)

default_button = ctk.CTkButton(
    temp_frame, 
    text="Restore Default Desktop", 
    command=reset_desktop_to_default,
    fg_color=SUCCESS_COLOR,
    hover_color="#1d7d74",
    font=ctk.CTkFont(size=14),
    width=200,
    height=40
)
default_button.pack(side="left", padx=5)

# Section for create a profile
create_frame = ctk.CTkFrame(main_frame)
create_frame.pack(fill="x", padx=10, pady=10)

ctk.CTkLabel(create_frame, text="Create new profile:", font=ctk.CTkFont(size=16)).pack(side="left", padx=10)
profile_entry = ctk.CTkEntry(create_frame, width=200)
profile_entry.pack(side="left", padx=10)

create_button = ctk.CTkButton(
    create_frame, 
    text="Create Profile", 
    command=create_profile,
    fg_color=SUCCESS_COLOR,
    hover_color="#1d7d74",
    font=ctk.CTkFont(size=14)
)
create_button.pack(side="left", padx=10)

# Status
status_frame = ctk.CTkFrame(root, fg_color="transparent")
status_frame.pack(fill="x", padx=10, pady=5)

status_label = ctk.CTkLabel(status_frame, text="", font=ctk.CTkFont(size=12))
status_label.pack(side="left", padx=10)

version_label = ctk.CTkLabel(status_frame, text="v1.0", font=ctk.CTkFont(size=12))
version_label.pack(side="right", padx=10)

# Check if there is a temp profile at the start
active_profile = check_temp_profile()
profile_var.set(active_profile)

context_menu_frame = ctk.CTkFrame(main_frame)
context_menu_frame.pack(fill="x", padx=10, pady=10)
register_button = ctk.CTkButton(
    context_menu_frame, 
    text="add right click shortcut", 
    command=register_context_menu,
    fg_color=PRIMARY_COLOR,
    hover_color=ACCENT_COLOR,
    font=ctk.CTkFont(size=14)
)
register_button.pack(side="left", padx=10)

unregister_button = ctk.CTkButton(
    context_menu_frame, 
    text="remove shortcut", 
    command=unregister_context_menu,
    fg_color=DANGER_COLOR,
    hover_color="#c1121f",
    font=ctk.CTkFont(size=14)
)
unregister_button.pack(side="left", padx=10)

# set start profile
update_profiles()

# setup the tray
tray_icon = setup_tray_icon()
def run_tray():
    tray_icon.run()

# start the tray thread
tray_thread = threading.Thread(target=run_tray, daemon=True)
tray_thread.start()

if __name__ == "__main__":
    try:
        root.withdraw()
        root.mainloop()
    except KeyboardInterrupt:
        if tray_icon:
            tray_icon.stop()
    finally:
        if tray_icon:
            tray_icon.stop()
        sys.exit(0)