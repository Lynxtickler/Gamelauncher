import os
import sys
import shutil
import win32com.client
import configparser
import tkinter as tk
import re
from subprocess import run as cmdcall
from subprocess import STARTUPINFO, STARTF_USESHOWWINDOW, SW_HIDE
from PIL import Image, ImageTk
from tkinter import messagebox, filedialog
from functools import partial
from ast import literal_eval


SCRIPT_PATH = os.path.dirname(os.path.abspath(__file__))
ICONS_PATH = os.path.join(SCRIPT_PATH, 'icons')
LINKS_PATH = os.path.join(SCRIPT_PATH, 'shortcuts')
DATAFILE_PATH = os.path.join(SCRIPT_PATH, 'targets.dat')
user_desktop_path = ''

SELECT_DESKTOP_TITLE = 'Select the path to your desktop'
ADD_GAME_TITLE = 'Select the game you want to add'

DEFAULT_WIDTH = 100
TOP_HEIGHT = DEFAULT_WIDTH * 0.85
BORDER_WIDTH = DEFAULT_WIDTH * 0.025
SECTION_PADDING = DEFAULT_WIDTH - TOP_HEIGHT

LABEL_PADDING = 4
FONT = 'ubuntu'
HEADER_FONT_SIZE = 8
LABEL_FONT_SIZE = 10
HEADER_X = 4
HEADER_OFFSET = -4
ODD_KEYS = {'Esc': DEFAULT_WIDTH * 1.1, 'Tab': DEFAULT_WIDTH * 1.45, 'Backspace': DEFAULT_WIDTH * 2, 'Ins': DEFAULT_WIDTH * 1.55,
            'Caps': DEFAULT_WIDTH * 1.7, 'Del': DEFAULT_WIDTH * 1.3, 'LShift': DEFAULT_WIDTH * 1.3, 'RShift': DEFAULT_WIDTH * 2.72}

KEYB_ROWS = (('Esc', 'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12'),
             ('§', '1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '+', '´', 'Backspace'),
             ('Tab', 'q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', 'å', '¨', 'Ins'),
             ('Caps', 'a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', 'ö', 'ä', "'", 'Del'),
             ('LShift', '<', 'z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.', '-', 'RShift'))
DUPLICATE_KEYS = ('LShift', 'RShift')
FORBIDDEN_KEYS = ('Esc', 'F4', '´', '¨', 'PgUp', 'Caps', 'PgDn', 'LShift', 'RShift', 'Ins', 'Del', 'Backspace')
SPECIAL_KEYS = {'Esc': '<Escape>', 'Backspace': '<BackSpace>', 'Tab': '<Tab>', 'å': '<aring>', 'Ins': '<Insert>',
                'ö': '<odiaeresis>', 'ä': '<adiaeresis>', 'Del': '<Delete>', '<': '<less>'}

COL_DARK = '#212121'
COL_LIGHT = '#f5f5f5'
COL_WHITE = '#ffffff'
COL_TEXT = '#ff4081'
COL_TEXT2 = '#757575'
COL_ADD = '#4caf50'
COL_DEL = '#9a0007'
COL_ADMIN = '#ffc107'
COL_NOADMIN = '#1976d2'
COL_DIV = '#616161'

ADD_MODE = 1
DELETE_MODE = 2
TOGGLE_ADMIN_MODE = 3


class Key:
    """
    GUI key object
    """
    def __init__(self, text, parent, button=None, label=None, game=None, admin=None, fun=None):
        self.text = text
        self.parent = parent
        self.button = button
        self.label = label
        self.game = game
        self.admin = admin
        self.fun = fun

    def launch(self, e=None):
        """
        Launch the shortcut bound to this key
        e: tkinter Event object, never needed
        """
        shellrun = os.path.join(SCRIPT_PATH, 'shellrun.exe')
        if self.game:
            link_path = os.path.join(LINKS_PATH, self.game + '.lnk')
            if not os.path.isfile(link_path):
                link_path = os.path.join(LINKS_PATH, self.game + '.url')
                if not os.path.isfile(link_path):
                    self.parent.pop_error('Shortcut is of unknown format.')
            if self.admin:
                print('running as admin')
                try:
                    elevator = os.path.join(SCRIPT_PATH, 'elevate.exe')
                    startupinfo = STARTUPINFO()
                    startupinfo.dwFlags = STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = SW_HIDE
                    cmdcall(f'"{elevator}" "{link_path}"', startupinfo=startupinfo)
                except Exception as e:
                    self.parent.pop_error(f'Error launching the game as admin:\n{e}')
                    return
            else:
                try:
                    # Using the AutoHotkey-generated utility because it was easier for me
                    # TODO: make pure Python
                    cmdcall(f'"{shellrun}" "{link_path}"')
                except Exception as e:
                    self.parent.pop_error(f'Error launching game:\n{e}')
            sys.exit()

    def bind_mousehover(self, message=None):
        """
        Bind this key to change the tip header on the GUI
        """
        if self.game is None:
            if message:
                self.button.bind('<Enter>', partial(self.parent.change_header_tip, game=message))
                self.button.bind('<Leave>', partial(self.parent.change_header_tip, game=''))
                self.label.bind('<Enter>', partial(self.parent.change_header_tip, game=message))
                return
            self.button.unbind('<Enter>')
            self.button.unbind('<Leave>')
            self.label.unbind('<Enter>')
        else:
            self.button.bind('<Enter>', partial(self.parent.change_header_tip, game=self.game))
            self.button.bind('<Leave>', partial(self.parent.change_header_tip, game=''))
            self.label.bind('<Enter>', partial(self.parent.change_header_tip, game=self.game))


class App:
    def __init__(self, master):
        """
        Read datafile, assign globals and instance vars and create the GUI
        master: Tkinter root
        """

        # Initialise class variables
        data = self.read_data()
        self.games = data[1]
        self.keys = {}
        global user_desktop_path
        user_desktop_path = data[0]
        desktop_changed = False
        while True:
            if user_desktop_path:
                break
            desktop_changed = True
            master.withdraw()
            answer = filedialog.askdirectory(title=SELECT_DESKTOP_TITLE)
            if answer:
                user_desktop_path = answer.replace('\\', '/')
                break
        if desktop_changed:
            if not self.write_data() == 0:
                self.pop_error('Data file could not be written to')
        # Configure window appearance
        master.overrideredirect(True)
        frame_width = 13 * (DEFAULT_WIDTH + BORDER_WIDTH) + ODD_KEYS['Backspace'] + 2 * BORDER_WIDTH
        frame_height = 5 * (DEFAULT_WIDTH + BORDER_WIDTH) + BORDER_WIDTH
        x = 1920 / 2 - frame_width / 2
        y = 1080 / 2 - frame_height / 2
        self.frame = tk.Frame(master, width=frame_width, height=frame_height, borderwidth=BORDER_WIDTH, relief=tk.RIDGE, bg=COL_DARK)
        self.header_tip = tk.Label(self.frame, fg=COL_TEXT, bg=COL_DARK, font=(FONT, HEADER_FONT_SIZE), text='')
        self.header_tip.place(x=HEADER_X, y=TOP_HEIGHT + HEADER_OFFSET)
        master.geometry('%dx%d+%d+%d' % (frame_width, frame_height, x, y))
        self.frame.focus_set()
        self.frame.pack_propagate(False)
        self.create_buttons()
        self.frame.pack()
        if len(sys.argv) > 1 and sys.argv[1] == '-i':
            # running the script this far (no windows will be shown) beforehand speeds up the launch next time from my experience
            sys.exit()

    def create_buttons(self):
        """
        Initialise keys and bind all that have assigned games
        """
        for i, row in enumerate(KEYB_ROWS):
            padding = 0
            for j, key in enumerate(row):
                this_width = ODD_KEYS[key] if (key in ODD_KEYS) else DEFAULT_WIDTH
                button = tk.Button(self.frame)
                button.config(bg=COL_DARK, activebackground=COL_DARK, highlightbackground=COL_DIV)
                button.place(x=padding, y=(DEFAULT_WIDTH + BORDER_WIDTH) * i, width=this_width, height=TOP_HEIGHT if i == 0 else DEFAULT_WIDTH)
                label = tk.Label(self.frame, font=(FONT, LABEL_FONT_SIZE), text=key[1::] if key in DUPLICATE_KEYS else key)
                label.config(fg=COL_TEXT, bg=COL_DARK, disabledforeground=COL_TEXT2)
                label.place(x=padding + LABEL_PADDING, y=(DEFAULT_WIDTH + BORDER_WIDTH) * i + LABEL_PADDING)
                keyobj = Key(key, self, button, label)
                self.keys[key] = keyobj
                try:
                    if key in self.games:
                        keyobj.game = self.games[key][0]
                        keyobj.admin = self.games[key][1]
                        keyobj.fun = keyobj.launch
                        self.assign_key_fun(keyobj)
                        self.assign_key_ico(keyobj)
                    else:
                        self.assign_key_fun(keyobj)
                except TypeError as e:
                    print(e)
                padding += this_width + BORDER_WIDTH
                if i == 0:
                    if j % 4 == 0:
                        padding += DEFAULT_WIDTH / 2
                    if key == 'Esc':
                        padding += DEFAULT_WIDTH * 0.43

    def assign_keys(self, mode=None, game=None):
        """
        Reassign all keys along the objects' attributes
        mode: Whether to assign keys as default, for adding a game or deleting one. If omitted or evaluates to False
            in place of a condition, default behaviour is used
        game: Game name without file extension when mode=ADD_MODE is used
        """
        if not mode:
            for i, key in self.keys.items():
                self.assign_key_fun(key)
                key.label.config(fg=COL_TEXT)
        elif mode == ADD_MODE:
            for i, key in self.keys.items():
                try:
                    if key.text == 'Esc':
                        self.assign_key_fun(key, self.assign_keys)
                        key.bind_mousehover(message='Cancel')
                    elif key.text not in FORBIDDEN_KEYS:
                        self.assign_key_fun(key, self.finish_adding, game, key)
                        key.label.config(fg=COL_ADD)
                except tk.TclError as e:
                    print(e)
                    continue
        elif mode == DELETE_MODE:
            self.change_header_tip(None, game='Choose the game you want to remove')
            for i, key in self.keys.items():
                try:
                    if key.text == 'Esc':
                        self.assign_key_fun(key, self.assign_keys)
                        key.bind_mousehover(message='Cancel')
                    elif key.text not in FORBIDDEN_KEYS and key.game:
                        self.assign_key_fun(key, self.finish_deleting, key)
                        key.label.config(fg=COL_DEL)
                except tk.TclError as e:
                    print(e)
                    continue
        elif mode == TOGGLE_ADMIN_MODE:
            self.change_header_tip(None, game='Choose a game you want to toggle admin rights for')
            for i, key in self.keys.items():
                try:
                    if key.text in ('Esc', 'Backspace'):
                        self.assign_key_fun(key, self.assign_keys)
                        key.bind_mousehover(message='Cancel')
                    elif key.text not in FORBIDDEN_KEYS and key.game:
                        self.assign_key_fun(key, self.finish_toggling, key)
                        if key.admin == 0:
                            key.label.config(fg=COL_ADMIN)
                        elif key.admin == 1:
                            key.label.config(fg=COL_NOADMIN)
                except tk.TclError as e:
                    print(e)
                    continue

    def assign_key_fun(self, key, callback_func=False, *args):
        """
        Assign a function to buttonpresses and keypresses
        key: Key object
        callback_func: Function to assign to key
        *args: Arbitrary number of args to be passed to the callback if it's defined
        """
        bind_tag = key.text if key.text not in SPECIAL_KEYS else SPECIAL_KEYS[key.text]
        key.bind_mousehover()
        if re.match('F\d\d?', key.text):
            bind_tag = '<F{}>'.format(key.text[1::])
        if key.game:
            key.fun = key.launch
        elif key.text == 'Esc':
            key.fun = self.close
            bind_tag = SPECIAL_KEYS['Esc']
            key.bind_mousehover(message='Close app')
        elif key.text == 'Ins':
            key.fun = self.add_game
            bind_tag = SPECIAL_KEYS['Ins']
            key.bind_mousehover(message='Add a game')
        elif key.text == 'Del':
            key.fun = self.delete_game
            bind_tag = SPECIAL_KEYS['Del']
            key.bind_mousehover(message='Delete a game')
        elif key.text == 'Backspace':
            key.fun = self.toggle_admin
            bind_tag = SPECIAL_KEYS['Backspace']
            key.bind_mousehover('Toggle admin rights for games')
        elif not callback_func:
            key.button.config(state=tk.DISABLED, command=None)
            key.label.config(state=tk.DISABLED)
            key.label.unbind('<Button-1>')
            self.frame.unbind(bind_tag)
            return

        if callback_func:
            key.fun = partial(callback_func, *args)
        key.button.config(state=tk.NORMAL, command=key.fun)
        key.label.config(state=tk.NORMAL)
        key.label.bind('<Button-1>', self.run_bind)
        self.frame.bind(bind_tag, self.run_bind)

    def assign_key_ico(self, key):
        """
        Set an icon displayed on the button. Use placeholder instead when unsuccessful
        key: The target key
        """
        if key.text in FORBIDDEN_KEYS:
            return False
        try:
            game_name = key.game
        except KeyError:
            # set game_name to an arbitrary non-None value that can't exist in any legal game name (a file path)
            game_name = '?null'
        icon_path = os.path.join(ICONS_PATH, game_name + '.ico')
        if not os.path.isfile(icon_path):
            icon_path = os.path.join(SCRIPT_PATH, 'placeholder.png')
            if not game_name == '?null':
                try:
                    if self.save_icon(game_name):
                        icon_path = os.path.join(ICONS_PATH, game_name + '.ico')
                except OSError as e:
                    self.pop_error(e)
        try:
            icon_image = ImageTk.PhotoImage(Image.open(icon_path))
        except FileNotFoundError:
            print(f'No image located at {icon_path}.')
            return False
        key.button.config(image=icon_image)
        key.button.image = icon_image
        return True

    # Hanging around in case the other implementation turns out to be shitty at any point
    """
    def save_ico(game_name):
        with winshell.shortcut(os.path.join(LINKS_PATH, game_name+'.lnk')) as link:
            print('shortcut location: {}'.format(os.path.join(LINKS_PATH, game_name+'.lnk')))
            print(link.icon_location)
            #get_icon(link.icon_location[0])
            ico_path = os.path.join(ICONS_PATH, game_name+'.ico')
            returncode = os.system('iconsaver.bat ' + '"'+link.icon_location[0]+'" ' + '"'+ico_path+'"')
            return returncode
    """

    def save_icon(self, game_name):
        """
        Extract icon from .lnk or .url file in steps until success
        game_name: Game name without extension
        """
        target = os.path.join(LINKS_PATH, game_name + '.lnk')
        if not os.path.isfile(target):
            target = os.path.join(LINKS_PATH, game_name + '.url')
        i = 0
        while True:
            i += 1
            targetlc = target.lower()
            if not target:
                self.pop_error(f'Could not extract icon. Add icon manually to {os.path.join(ICONS_PATH, game_name)}.ico')
                return False
            elif targetlc.endswith('.ico'):
                new_ico = os.path.join(ICONS_PATH, game_name + '.ico')
                shutil.copy(target, new_ico)
                Image.open(new_ico).resize((32, 32), Image.ANTIALIAS).save(new_ico)
                break
            elif targetlc.endswith('.exe'):
                ico_path = os.path.join(ICONS_PATH, game_name + '.ico')
                iconsaver_command = '"{}" "{}" "{}"'.format(os.path.join(SCRIPT_PATH, 'iconsaver.bat'), target, ico_path)
                # iconsaver.bat is a batch file copied straight from stackoverflow or superuser or so and it works really damn well
                cmdcall(iconsaver_command, shell=True)
                break
            elif targetlc.endswith('.lnk'):
                shell = win32com.client.Dispatch('WScript.Shell')
                shlink = shell.CreateShortCut(target)
                if shlink.IconLocation:
                    target = shlink.IconLocation.split(',')[0]
                    if not os.path.isfile(target):
                        target = shlink.Targetpath
                else:
                    target = shlink.Targetpath
            elif targetlc.endswith('.url'):
                config = configparser.ConfigParser()
                config.read(target)
                print(config['InternetShortcut']['url'].split('/'))
                if len(config['InternetShortcut']['url'].split('/')[-1]) > 7:  # prevents problems with non-steam games added to steam
                    target = config['InternetShortcut']['iconfile']
                    continue
                new_ico = os.path.join(ICONS_PATH, game_name + '.ico')
                shutil.copy(config['InternetShortcut']['iconfile'], new_ico)
                Image.open(new_ico).resize((32, 32), Image.ANTIALIAS).save(new_ico)
                break
            else:
                self.pop_error('Unhandled filetype: "{}" occured while binding {}'.format(os.path.splitext(target), game_name))
                target = False
        return True

    # handler functions (add default None to all events)
    # --------------------------------------------------

    def close(self, event=None):
        """
        Close app
        event: In case of a kb event this must be retrieved but it's always ignored
        """
        sys.exit()

    def run_bind(self, event):
        """
        Run the function bound to the key
        event: Event object that's handled if it's either type KeyPress (2, from keyboard) or ButtonPress (4, from GUI)
        """
        if int(event.type) == 4:  # the only realiable way to get the event type
            key_text = event.widget.cget('text')
            bind_tag = key_text if key_text not in SPECIAL_KEYS else SPECIAL_KEYS[key_text]
            self.keys[bind_tag].fun()
            return
        if int(event.type) == 2:
            bind_tag = event.keysym
            exception_keys = {'Escape': 'Esc', 'section': '§', 'plus': '+', 'BackSpace': 'Backspace', 'aring': 'å', 'odiaeresis': 'ö',
                              'adiaeresis': 'ä', 'quoteright': "'", 'Shift_L': 'Shift', 'less': '<', 'comma': ',', 'period': '.',
                              'minus': '-', 'Shift_R': 'Shift', 'Insert': 'Ins', 'Delete': 'Del'}
            if event.keysym in exception_keys:
                bind_tag = exception_keys[event.keysym]
            self.keys[bind_tag].fun()

        # Not needed, left to hang around in case of a strong will of adding new features
        """
        if event.char in multi_keys.values():
            print('was space after multikey')
            #key[event.char].fun()
            return
        if event.keysym == 'Multi_key':
            print('was multikey')
            agpress('space')
            return
            try:
                keys[multi_keys[event.keycode]].fun()
            except KeyError:
                pass
            return
        """

    # --------------------------------------------------

    def add_game(self):
        """
        Ask to choose a shortcut from desktop and call assign_keys to reassign keys so that when they are pressed the next time,
        the game in question is assigned to that key.
        """
        game_path = filedialog.askopenfilename(initialdir='Desktop', filetypes=(('Shortcuts', '*.lnk *.url'), ('All files', '*.*')), title=ADD_GAME_TITLE)
        print('-----Closed shortcut picker dialog-----')  # without any action being taken the app hangs for several seconds after closing the previous dialog
        if not (game_path.endswith('.lnk') or game_path.endswith('.url')):
            return False
        self.assign_keys(ADD_MODE, game_path)
        self.change_header_tip(None, game='Choose what key to bind the game to')

    def finish_adding(self, game, key=None):
        """
        Finish the game adding process by moving the shortcut to link folder, assigning the key and updating data structures in memory and disk
        game: Path to desired game
        key: The key to bind the game to. Can be omitted in case game is '?' because key is then never needed
        """
        if not game:
            return
        self.assign_keys()
        if not key:
            return
        if game == '?':
            # Adding is cancelled, a.k.a esc was pressed
            return
        dest_path = os.path.join(LINKS_PATH, os.path.basename(game))
        try:
            shutil.move(game, dest_path)
        except IOError as e:
            print(e)
            self.pop_error('Could not save shortcut from desktop')
        basename = os.path.basename(dest_path)
        key.game = basename.replace(os.path.splitext(basename)[1], '')
        key.admin = 0
        self.assign_key_fun(key)
        data = self.read_data()
        self.games = data[1]
        self.games.update({key.text: [key.game, key.admin]})
        if not self.write_data() == 0:
            self.pop_error('Data file could not be written to')
        success = self.assign_key_ico(key)
        if not success:
            self.pop_error('Could not extract proper icon from the shortcut')

    def delete_game(self):
        """Oneliner for consistent method names for adding and deleting games"""
        self.assign_keys(DELETE_MODE)

    def finish_deleting(self, key):
        """
        Finish the game deletion process by moving the shortcut from link folder to desktop, deleting its .ico, unassigning the key
        and updating data structures in memory and disk
        key: The key to unbind
        """
        del self.games[key.text]
        if not self.write_data() == 0:
            self.pop_error('Data file could not be written to')
        try:
            lnk_path = os.path.join(LINKS_PATH, key.game + '.lnk')
            url_path = os.path.join(LINKS_PATH, key.game + '.url')
            shutil.move(lnk_path if os.path.isfile(lnk_path) else url_path, user_desktop_path)
            os.remove(os.path.join(ICONS_PATH, key.game + '.ico'))
        except shutil.Error:
            print('Shortcut was already on desktop')
        except FileNotFoundError:
            pass
        key.button.image = None
        key.button.config(image='')
        key.game = None
        key.fun = None
        key.bind_mousehover()
        self.assign_key_fun(key)
        self.assign_keys()

    def toggle_admin(self):
        """Oneliner for consistent method names for adding and deleting games"""
        self.assign_keys(TOGGLE_ADMIN_MODE)

    def finish_toggling(self, key):
        """
        Change a key's admin property between 0 and 1.
        key: Target key object
        """
        if key.admin == 0:
            self.games[key.text][1] = 1
            key.admin = 1
            key.label.config(fg=COL_ADD)
        else:
            self.games[key.text][1] = 0
            key.admin = 0
            key.label.config(fg=COL_DEL)
        if not self.write_data() == 0:
            self.pop_error('Data file could not be written to')
        self.assign_keys()

    def write_data(self):
        """
        Write data into the datafile with formatting
        """
        try:
            with open(DATAFILE_PATH, 'w', encoding='utf-8') as f:
                print(f'("{user_desktop_path}",', file=f)
                print('{', file=f)
                for i, (key, (game, admin)) in enumerate(self.games.items()):
                    print(f'    "{key}": ["{game}", {admin}]', file=f, end='')
                    if not i == len(self.games) - 1:
                        print(',', file=f)
                print('\n})', file=f)
            return 0
        except Exception as e:
            print(e)
            return 1

    def read_data(self):
        """
        Read data in datafile. Return the (str, dict) tuple with the read data or both datatypes but empty in case of error or corrupt file
        """
        with open(DATAFILE_PATH, 'r', encoding='utf-8') as f:
            data = f.read()
        if data is None:
            return ('', {})
        try:
            return literal_eval(data)
        except Exception as e:
            print(e)
            self.pop_error('Data file is corrupted')
            return ('', {})

    def pop_error(self, error):
        """
        Pop error with some text
        error: Text to display
        """
        messagebox.showerror('Error', error)

    def change_header_tip(self, event, game=''):
        """
        Change the text in header label
        event: Will be ignored
        game: Message to display, often a game name
        """
        self.header_tip.config(text=game)

    def change_desktop_path(self, master):
        """
        Change the desktop path. Never called anywhere yet though
        master: Tkinter root
        """
        master.withdraw()
        answer = filedialog.askdirectory(title=SELECT_DESKTOP_TITLE)
        if answer:
            global user_desktop_path
            user_desktop_path = answer.replace('\\', '/')
            self.write_data()


if __name__ == "__main__":
    root = tk.Tk()
    if '-h' in sys.argv:
        root.withdraw()
        messagebox.showinfo(f'{os.path.basename(__file__)} - Help', 'Use param -i for initialising GUI on boot and use no params for normal behaviour')
        sys.exit()
    app = App(root)
    root.mainloop()
