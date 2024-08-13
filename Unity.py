import os
import asyncio
import json
import pickle
from datetime import datetime
from tkinter import Tk, Listbox, Entry, Label, Scrollbar, StringVar, Toplevel, BooleanVar, Frame, Text, Button, filedialog, simpledialog, messagebox, Checkbutton
from tkinter.ttk import Style, Combobox, Entry as TtkEntry, Button as TtkButton, Label as TtkLabel, Checkbutton as TtkCheckbutton, Frame as TtkFrame
from tkinter.constants import LEFT, BOTH, RIGHT, Y
from pynput import keyboard
from pywinauto import Desktop
import pygetwindow as gw
import pyautogui
from PIL import Image, ImageTk
import threading
import queue
from ttkthemes import ThemedTk

# Create directory if it doesn't exist
if not os.path.exists('Formation'):
    os.makedirs('Formation')

# Initialize Tkinter with ThemedTk
root = ThemedTk(theme="equilux")  # Replace "equilux" with any dark theme you like from ttkthemes
root.title("Dofus Manager")

icon_path = os.path.join(os.path.dirname(__file__), 'icons', 'dofus.ico')
root.iconbitmap(icon_path)

# Set background color to match the theme
root.configure(bg="#2b2b2b")  # Adjust the color as needed to match your theme

# Apply a custom style
style = Style(root)
style.theme_use("equilux")  # Replace "equilux" with any dark theme you like from ttkthemes

# Configure ttk styles
style.configure('TLabel', background='#2b2b2b', foreground='#f0f0f0')
style.configure('TFrame', background='#2b2b2b')
style.configure('TEntry', fieldbackground='#3e3e3e', foreground='#f0f0f0', background='#2b2b2b')
style.configure('TButton', background='#3e3e3e', foreground='#f0f0f0')

# Global variables
current_windows = []
ignored_windows = []
is_searching = False
windows_lock = asyncio.Lock()
last_update_time = StringVar()
auto_refresh_enabled = BooleanVar(value=True)
current_order_name = StringVar(value="Aucun ordre chargé")
search_var = StringVar()
status_var = StringVar()
cached_windows = {}
task_queue = queue.Queue()

# Key binding variables
forward_key = StringVar(value='f5')
backward_key = StringVar(value='f6')
ignore_key = StringVar(value='f7')

def update_status(message):
    print(message)  # Print to console for debugging
    status_var.set(message)
    log_text.insert('end', f"{message}\n")
    log_text.see('end')
    if int(log_text.index('end-1c').split('.')[0]) > 100:
        log_text.delete(1.0, 2.0)

def search_windows(event):
    query = search_var.get().lower()
    listbox.delete(0, 'end')
    for window in current_windows:
        if query in window[1].lower():
            listbox.insert('end', f"{window[1]} ({window[2]})")

def bind_shortcuts():
    root.bind('<Control-r>', lambda event: task_queue.put(('update_windows',)))
    root.bind('<Control-s>', lambda event: task_queue.put(('save_order',)))
    root.bind('<Control-l>', lambda event: task_queue.put(('load_order', event)))
    root.bind('<Control-q>', lambda event: root.quit())

bind_shortcuts()

async def update_windows():
    global current_windows, ignored_windows
    async with windows_lock:
        try:
            desktop = Desktop(backend="uia")
            windows = desktop.windows(class_name="UnityWndClass")
            dofus_windows = [window for window in windows if "Dofus" in window.window_text()]
            new_windows = [(window, window.window_text(), str(window.handle)) for window in dofus_windows]
            current_windows = list(dict.fromkeys(new_windows))
            update_status(f"Found {len(current_windows)} Dofus windows.")
            task_queue.put(('update_listbox', current_windows, ignored_windows))
            last_update_time.set(f"Dernière mise à jour: {datetime.now().strftime('%H:%M:%S')}")
            update_status(f"Liste des fenêtres mise à jour à {datetime.now().strftime('%H:%M:%S')}")
        except Exception as e:
            update_status(f"Erreur lors de la mise à jour des fenêtres : {e}")

async def auto_refresh():
    while True:
        if auto_refresh_enabled.get():
            update_status("Auto-refreshing windows...")
            task_queue.put(('update_windows',))
        await asyncio.sleep(10)

def start_auto_refresh():
    print("Starting auto-refresh...")
    asyncio.create_task(auto_refresh())

def ignore_window_by_handle(window_handle):
    for window in current_windows:
        if window[2] == window_handle:
            ignored_windows.append(window)
            current_windows.remove(window)
            update_listbox_ui(current_windows, ignored_windows)
            return True
    return False

def unignore_window_by_handle(window_handle):
    for window in ignored_windows:
        if window[2] == window_handle:
            current_windows.append(window)
            ignored_windows.remove(window)
            update_listbox_ui(current_windows, ignored_windows)
            return True
    return False

def toggle_ignore_window():
    try:
        active_window = gw.getActiveWindow()
        if active_window is None:
            update_status("No active window found.")
            return
        window_handle = str(active_window._hWnd)
        if not ignore_window_by_handle(window_handle):
            unignore_window_by_handle(window_handle)
    except Exception as e:
        update_status(f"Error toggling ignore state: {e}")

def ignore_window():
    selection = listbox.curselection()
    if selection:
        selected_item = listbox.get(selection)
        for tuple in current_windows:
            window = tuple[0]
            if window.window_text() == selected_item.split(" (")[0]:
                ignored_windows.append(tuple)
                current_windows.remove(tuple)
                break
        listbox.delete(selection)
        ignored_listbox.insert('end', selected_item)

def unignore_window():
    selection = ignored_listbox.curselection()
    if selection:
        selected_item = ignored_listbox.get(selection)
        for tuple in ignored_windows:
            window = tuple[0]
            if window.window_text() == selected_item.split(" (")[0]:
                current_windows.append(tuple)
                ignored_windows.remove(tuple)
                break
        ignored_listbox.delete(selection)
        listbox.insert('end', selected_item)

def move_up():
    selection = listbox.curselection()
    if selection:
        index = selection[0]
        if index > 0:
            listbox.insert(index-1, listbox.get(index))
            listbox.delete(index+1)
            listbox.selection_set(index-1)
    refresh_order()

def move_down():
    selection = listbox.curselection()
    if selection:
        index = selection[0]
        if index < listbox.size()-1:
            listbox.insert(index+2, listbox.get(index))
            listbox.delete(index)
            listbox.selection_set(index+1)
    refresh_order()

def refresh_order():
    global current_windows
    ordered_list = []
    for item in listbox.get(0, "end"):
        window_name = item.split(" (")[0]
        for window in current_windows:
            if window[1] == window_name:
                ordered_list.append(window)
                break
    current_windows = ordered_list
    update_status(f"Order refreshed: {ordered_list}")

def save_order():
    window_order = [item.split(" ")[0] for item in listbox.get(0, 'end')]
    order_name = simpledialog.askstring("Nom de l'ordre", "Veuillez entrer un nom pour l'ordre")
    if order_name:
        try:
            with open(f'Formation/{order_name}.pkl', 'wb') as f:
                pickle.dump(window_order, f)
            saved_orders.append(order_name)
            saved_orders.sort()
            update_order_listbox()
            messagebox.showinfo("Information", f"L'ordre '{order_name}' a été sauvegardé.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {e}")
    else:
        messagebox.showerror("Erreur", "Veuillez entrer un nom pour l'ordre.")

def rename_order():
    selection = order_listbox.curselection()
    if selection:
        old_name = order_listbox.get(selection[0])
        new_name = simpledialog.askstring("Renommer l'ordre", "Veuillez entrer un nouveau nom pour l'ordre")
        if new_name:
            old_path = f'Formation/{old_name}.pkl'
            new_path = f'Formation/{new_name}.pkl'
            if os.path.exists(new_path):
                messagebox.showerror("Erreur", "Un ordre avec ce nom existe déjà.")
            else:
                os.rename(old_path, new_path)
                saved_orders.remove(old_name)
                saved_orders.append(new_name)
                saved_orders.sort()
                update_order_listbox()
                messagebox.showinfo("Information", f"L'ordre '{old_name}' a été renommé en '{new_name}'.")
        else:
            messagebox.showerror("Erreur", "Veuillez entrer un nouveau nom pour l'ordre.")

def clean_up_duplicates():
    global current_windows
    unique_windows = {}
    for window in current_windows:
        if window[2] not in unique_windows:
            unique_windows[window[2]] = window
    current_windows = list(unique_windows.values())

def update_listbox_ui(new_windows, new_ignored_windows):
    listbox.delete(0, 'end')
    for window in new_windows:
        window_text = f"{window[1]} ({window[2]})"
        if window[1]:  # Ensure window has a name
            listbox.insert('end', window_text)
    ignored_listbox.delete(0, 'end')
    for window in new_ignored_windows:
        window_text = f"{window[1]} ({window[2]})"
        ignored_listbox.insert('end', window_text)

async def load_order(event=None):
    selection = order_listbox.curselection()
    if selection:
        order_name = order_listbox.get(selection[0])
        try:
            with open(f'Formation/{order_name}.pkl', 'rb') as f:
                saved_character_names = pickle.load(f)
            update_status(f"Loading order: {order_name} with characters {saved_character_names}")
            await update_windows()
            await asyncio.sleep(1)
            listbox.delete(0, 'end')
            unique_windows = set()
            for name in saved_character_names:
                for window_tuple in current_windows:
                    if window_tuple[1].startswith(name):
                        if window_tuple[2] not in unique_windows:
                            window_text = f"{window_tuple[1]} ({window_tuple[2]})"
                            listbox.insert('end', window_text)
                            unique_windows.add(window_tuple[2])
                            update_status(f"Added {window_text} to listbox")
                        else:
                            update_status(f"Duplicate window detected and ignored: {window_tuple[1]} ({window_tuple[2]})")
            clean_up_duplicates()
            refresh_order()
            current_order_name.set(f"Ordre chargé: {order_name}")
            auto_refresh_enabled.set(False)
        except FileNotFoundError:
            messagebox.showerror("Erreur", f"L'ordre '{order_name}' n'a pas été trouvé.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du chargement : {e}")

def import_order():
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    if file_path:
        try:
            with open(file_path, 'r') as f:
                order = json.load(f)
            order_name = os.path.basename(file_path).replace('.json', '')
            with open(f'Formation/{order_name}.pkl', 'wb') as f:
                pickle.dump(order, f)
            saved_orders.append(order_name)
            saved_orders.sort()
            update_order_listbox()
            messagebox.showinfo("Information", f"L'ordre '{order_name}' a été importé.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'importation : {e}")

def export_order():
    window_order = [item.split(" ")[0] for item in listbox.get(0, 'end')]
    order_name = simpledialog.askstring("Nom de l'ordre", "Veuillez entrer un nom pour l'ordre")
    if order_name:
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")], initialfile=f"{order_name}.json")
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(window_order, f)
                messagebox.showinfo("Information", f"L'ordre '{order_name}' a été exporté.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'exportation : {e}")
    else:
        messagebox.showerror("Erreur", "Veuillez entrer un nom pour l'ordre.")

def create_tooltip(widget, text):
    tooltip = Toplevel(widget)
    tooltip.wm_overrideredirect(True)
    tooltip.wm_geometry("+0+0")
    label = Label(tooltip, text=text, relief='solid', borderwidth=1)
    label.pack()

    def enter(event):
        x = event.widget.winfo_rootx() + 20
        y = event.widget.winfo_rooty() + 20
        tooltip.wm_geometry(f"+{x}+{y}")
        tooltip.wm_deiconify()
    
    def leave(event):
        tooltip.wm_withdraw()
    
    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)
    tooltip.withdraw()

def setup_buttons():
    def resize_icon(icon_path, size=(20, 20)):
        img = Image.open(icon_path)
        img = img.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    refresh_icon = resize_icon('icons/refresh.png')
    ignore_icon = resize_icon('icons/ignore.png')
    unignore_icon = resize_icon('icons/unignore.png')
    up_icon = resize_icon('icons/up.png')
    down_icon = resize_icon('icons/down.png')
    save_icon = resize_icon('icons/save.png')
    rename_icon = resize_icon('icons/rename.png')
    import_icon = resize_icon('icons/import.png')
    export_icon = resize_icon('icons/export.png')

    button_frame = TtkFrame(root)
    button_frame.grid(row=1, column=3, rowspan=3, padx=10, pady=5, sticky='nsew')

    start_button = TtkButton(button_frame, text=" Rafraichir", image=refresh_icon, compound=LEFT, command=lambda: task_queue.put(('update_windows',)))
    start_button.grid(row=0, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(start_button, "Cliquez pour rafraîchir la liste des fenêtres Dofus.")

    ignore_button = TtkButton(button_frame, text=" Ignorer", image=ignore_icon, compound=LEFT, command=ignore_window)
    ignore_button.grid(row=1, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(ignore_button, "Cliquez pour ignorer la fenêtre sélectionnée et la déplacer dans la liste des fenêtres ignorées.")

    unignore_button = TtkButton(button_frame, text=" Ne plus Ignorer", image=unignore_icon, compound=LEFT, command=unignore_window)
    unignore_button.grid(row=2, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(unignore_button, "Cliquez pour ne plus ignorer la fenêtre sélectionnée et la replacer dans la liste principale.")

    move_up_button = TtkButton(button_frame, text=" Haut", image=up_icon, compound=LEFT, command=move_up)
    move_up_button.grid(row=3, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(move_up_button, "Cliquez pour déplacer la fenêtre sélectionnée vers le haut.")

    move_down_button = TtkButton(button_frame, text=" Bas", image=down_icon, compound=LEFT, command=move_down)
    move_down_button.grid(row=4, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(move_down_button, "Cliquez pour déplacer la fenêtre sélectionnée vers le bas.")

    save_order_button = TtkButton(button_frame, text=" Sauvegarder", image=save_icon, compound=LEFT, command=save_order)
    save_order_button.grid(row=5, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(save_order_button, "Cliquez pour sauvegarder l'ordre actuel des fenêtres.")

    rename_order_button = TtkButton(button_frame, text=" Renommer", image=rename_icon, compound=LEFT, command=rename_order)
    rename_order_button.grid(row=6, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(rename_order_button, "Cliquez pour renommer l'ordre sélectionné.")

    import_order_button = TtkButton(button_frame, text=" Importer", image=import_icon, compound=LEFT, command=import_order)
    import_order_button.grid(row=7, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(import_order_button, "Cliquez pour importer un ordre depuis un fichier JSON.")

    export_order_button = TtkButton(button_frame, text=" Exporter", image=export_icon, compound=LEFT, command=export_order)
    export_order_button.grid(row=8, column=0, padx=10, pady=5, sticky='ew')
    create_tooltip(export_order_button, "Cliquez pour exporter l'ordre actuel vers un fichier JSON.")

    start_button.image = refresh_icon
    ignore_button.image = ignore_icon
    unignore_button.image = unignore_icon
    move_up_button.image = up_icon
    move_down_button.image = down_icon
    save_order_button.image = save_icon
    rename_order_button.image = rename_icon
    import_order_button.image = import_icon
    export_order_button.image = export_icon

def update_order_listbox():
    order_listbox.delete(0, 'end')
    for order in saved_orders:
        order_listbox.insert('end', order)

def rotate_windows(direction):
    global current_windows
    if direction == 'forward' and current_windows:
        element = current_windows.pop(0)
        current_windows.append(element)
        update_status(f"Rotated to: {current_windows[0][1]} ({current_windows[0][2]})")
        current_windows[0][0].set_focus()
    elif direction == 'backward' and current_windows:
        element = current_windows.pop()
        current_windows.insert(0, element)
        update_status(f"Rotated to: {current_windows[0][1]} ({current_windows[0][2]})")
        current_windows[0][0].set_focus()

def on_press(key):
    global current_windows
    try:
        print(f"Key pressed: {key}")  # Debugging log
        if hasattr(key, 'char'):
            if key.char == forward_key.get():
                update_status(f"Rotating forward with key: {key.char}")
                rotate_windows('forward')
            elif key.char == backward_key.get():
                update_status(f"Rotating backward with key: {key.char}")
                rotate_windows('backward')
            elif key.char == ignore_key.get():
                update_status(f"Ignoring/Unignoring window with key: {key.char}")
                toggle_ignore_window()
        elif key == keyboard.Key[forward_key.get()]:
            update_status(f"Rotating forward with key: {key}")
            rotate_windows('forward')
        elif key == keyboard.Key[backward_key.get()]:
            update_status(f"Rotating backward with key: {key}")
            rotate_windows('backward')
        elif key == keyboard.Key[ignore_key.get()]:
            update_status(f"Ignoring/Unignoring window with key: {key}")
            toggle_ignore_window()
    except Exception as e:
        update_status(f"Error handling key press: {e}")

def open_key_binding_window():
    key_binding_window = Toplevel(root)
    key_binding_window.title("Configurer les touches")

    def set_key(key_var):
        def on_key_press(event):
            key_var.set(event.keysym.lower())
            key_binding_window.unbind('<KeyPress>')
        key_binding_window.bind('<KeyPress>', on_key_press)

    forward_key_label = TtkLabel(key_binding_window, text="Touche de rotation avant:")
    forward_key_label.pack(padx=10, pady=5, fill='x')
    forward_key_button = TtkButton(key_binding_window, textvariable=forward_key, command=lambda: set_key(forward_key))
    forward_key_button.pack(padx=10, pady=5, fill='x')

    backward_key_label = TtkLabel(key_binding_window, text="Touche de rotation arrière:")
    backward_key_label.pack(padx=10, pady=5, fill='x')
    backward_key_button = TtkButton(key_binding_window, textvariable=backward_key, command=lambda: set_key(backward_key))
    backward_key_button.pack(padx=10, pady=5, fill='x')

    ignore_key_label = TtkLabel(key_binding_window, text="Touche pour ignorer/ne plus ignorer une fenêtre:")
    ignore_key_label.pack(padx=10, pady=5, fill='x')
    ignore_key_button = TtkButton(key_binding_window, textvariable=ignore_key, command=lambda: set_key(ignore_key))
    ignore_key_button.pack(padx=10, pady=5, fill='x')

scrollbar = Scrollbar(root)
scrollbar.grid(row=1, column=2, rowspan=2, sticky='ns')

search_label = TtkLabel(root, text="Rechercher :")
search_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')

search_entry = TtkEntry(root, textvariable=search_var)
search_entry.grid(row=0, column=1, columnspan=3, padx=10, pady=5, sticky='ew')
search_entry.bind('<KeyRelease>', search_windows)

managed_frame = TtkFrame(root)
managed_frame.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

order_label = TtkLabel(managed_frame, textvariable=current_order_name)
order_label.pack(padx=10, pady=5, fill='x')

managed_label = TtkLabel(managed_frame, text="Fenêtres gérées")
managed_label.pack(padx=10, pady=5, fill='x')

listbox = Listbox(managed_frame, width=50, yscrollcommand=scrollbar.set, bg='#2b2b2b', fg='#f0f0f0')
listbox.pack(side=LEFT, fill=BOTH, expand=True)

ignored_frame = TtkFrame(root)
ignored_frame.grid(row=1, column=1, padx=10, pady=5, sticky='nsew')

ignored_label = TtkLabel(ignored_frame, text="Fenêtres ignorées")
ignored_label.pack(padx=10, pady=5, fill='x')

ignored_listbox = Listbox(ignored_frame, width=50, yscrollcommand=scrollbar.set, bg='#2b2b2b', fg='#f0f0f0')
ignored_listbox.pack(side=LEFT, fill=BOTH, expand=True)

order_frame = TtkFrame(root)
order_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky='nsew')

order_label = TtkLabel(order_frame, text="Ordres enregistrés")
order_label.pack(padx=10, pady=5, fill='x')

order_listbox = Listbox(order_frame, width=50, bg='#2b2b2b', fg='#f0f0f0')
order_listbox.pack(side=LEFT, fill=BOTH, expand=True)
order_listbox.bind('<<ListboxSelect>>', lambda event: task_queue.put(('load_order', event)))

saved_orders = sorted([f[:-4] for f in os.listdir('Formation') if f.endswith('.pkl')])
update_order_listbox()

log_frame = TtkFrame(root)
log_frame.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')

log_text = Text(log_frame, height=10, wrap='word', bg='#2b2b2b', fg='#f0f0f0')
log_text.pack(side=LEFT, fill=BOTH, expand=True)

log_scroll = Scrollbar(log_frame, command=log_text.yview)
log_scroll.pack(side=RIGHT, fill=Y)
log_text.config(yscrollcommand=log_scroll.set)

scrollbar.config(command=listbox.yview)
setup_buttons()

auto_refresh_checkbutton = TtkCheckbutton(root, text="Activer le rafraîchissement automatique", variable=auto_refresh_enabled)
auto_refresh_checkbutton.grid(row=3, column=2, padx=10, pady=5, sticky='ew')
create_tooltip(auto_refresh_checkbutton, "Cochez pour activer/désactiver le rafraîchissement automatique.")

last_update_label = TtkLabel(root, textvariable=last_update_time)
last_update_label.grid(row=4, column=2, padx=10, pady=5, sticky='ew')
last_update_time.set("Dernière mise à jour: jamais")

key_binding_button = TtkButton(root, text="Configurer les touches", command=open_key_binding_window)
key_binding_button.grid(row=4, column=0, padx=10, pady=5, sticky='ew')

# Drag and Drop functionality
def on_listbox_drag_start(event, listbox):
    listbox.drag_data = {"x": event.x, "y": event.y, "index": listbox.nearest(event.y)}

def on_listbox_drag_motion(event, listbox):
    listbox.drag_data["x"] = event.x
    listbox.drag_data["y"] = event.y

def on_listbox_drag_drop(event, listbox):
    index = listbox.nearest(event.y)
    if index != listbox.drag_data["index"]:
        item = listbox.get(listbox.drag_data["index"])
        listbox.delete(listbox.drag_data["index"])
        listbox.insert(index, item)
        if listbox == order_listbox:
            saved_orders.insert(index, saved_orders.pop(listbox.drag_data["index"]))
        else:
            refresh_order()

listbox.bind("<Button-1>", lambda event: on_listbox_drag_start(event, listbox))
listbox.bind("<B1-Motion>", lambda event: on_listbox_drag_motion(event, listbox))
listbox.bind("<ButtonRelease-1>", lambda event: on_listbox_drag_drop(event, listbox))

order_listbox.bind("<Button-1>", lambda event: on_listbox_drag_start(event, order_listbox))
order_listbox.bind("<B1-Motion>", lambda event: on_listbox_drag_motion(event, order_listbox))
order_listbox.bind("<ButtonRelease-1>", lambda event: on_listbox_drag_drop(event, order_listbox))

def process_tasks():
    while not task_queue.empty():
        task = task_queue.get()
        print(f"Processing task: {task}")  # Debugging log
        if task[0] == 'update_windows':
            try:
                asyncio.run_coroutine_threadsafe(update_windows(), loop)
            except RuntimeError as e:
                update_status(f"RuntimeError: {e}")
        elif task[0] == 'update_listbox':
            current_windows, ignored_windows = task[1], task[2]
            update_listbox_ui(current_windows, ignored_windows)
        elif task[0] == 'load_order':
            asyncio.run_coroutine_threadsafe(load_order(), loop)
        task_queue.task_done()
    root.after(100, process_tasks)

async def main_async():
    start_auto_refresh()
    while True:
        await asyncio.sleep(1)

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        root.quit()
        root.destroy()

if __name__ == '__main__':
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    threading.Thread(target=loop.run_forever, daemon=True).start()
    asyncio.run_coroutine_threadsafe(main_async(), loop)
    listener = keyboard.Listener(on_press=on_press)
    listener.start()
    root.after(100, process_tasks)
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()