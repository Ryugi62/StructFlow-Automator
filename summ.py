import sys
import win32gui
import win32process
import win32api
import win32con
import win32ui
import psutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QTreeWidget, QTreeWidgetItem, QVBoxLayout, QHBoxLayout, QWidget, QLineEdit, QPushButton, QTextEdit, QScrollArea, QLabel, QSpinBox
from PyQt5.QtCore import Qt, QTimer

def get_window_style(hwnd):
    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
    styles = []
    style_dict = {
        win32con.WS_BORDER: "WS_BORDER",
        win32con.WS_CAPTION: "WS_CAPTION",
        win32con.WS_CHILD: "WS_CHILD",
        win32con.WS_CLIPCHILDREN: "WS_CLIPCHILDREN",
        win32con.WS_CLIPSIBLINGS: "WS_CLIPSIBLINGS",
        win32con.WS_DISABLED: "WS_DISABLED",
        win32con.WS_DLGFRAME: "WS_DLGFRAME",
        win32con.WS_GROUP: "WS_GROUP",
        win32con.WS_HSCROLL: "WS_HSCROLL",
        win32con.WS_MAXIMIZE: "WS_MAXIMIZE",
        win32con.WS_MAXIMIZEBOX: "WS_MAXIMIZEBOX",
        win32con.WS_MINIMIZE: "WS_MINIMIZE",
        win32con.WS_MINIMIZEBOX: "WS_MINIMIZEBOX",
        win32con.WS_OVERLAPPED: "WS_OVERLAPPED",
        win32con.WS_POPUP: "WS_POPUP",
        win32con.WS_SIZEBOX: "WS_SIZEBOX",
        win32con.WS_SYSMENU: "WS_SYSMENU",
        win32con.WS_TABSTOP: "WS_TABSTOP",
        win32con.WS_THICKFRAME: "WS_THICKFRAME",
        win32con.WS_VISIBLE: "WS_VISIBLE",
        win32con.WS_VSCROLL: "WS_VSCROLL"
    }
    for s in style_dict:
        if style & s:
            styles.append(style_dict[s])
    return ', '.join(styles)

def get_window_ex_style(hwnd):
    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
    ex_styles = []
    ex_style_dict = {
        win32con.WS_EX_ACCEPTFILES: "WS_EX_ACCEPTFILES",
        win32con.WS_EX_APPWINDOW: "WS_EX_APPWINDOW",
        win32con.WS_EX_CLIENTEDGE: "WS_EX_CLIENTEDGE",
        win32con.WS_EX_COMPOSITED: "WS_EX_COMPOSITED",
        win32con.WS_EX_CONTEXTHELP: "WS_EX_CONTEXTHELP",
        win32con.WS_EX_CONTROLPARENT: "WS_EX_CONTROLPARENT",
        win32con.WS_EX_DLGMODALFRAME: "WS_EX_DLGMODALFRAME",
        win32con.WS_EX_LAYERED: "WS_EX_LAYERED",
        win32con.WS_EX_LAYOUTRTL: "WS_EX_LAYOUTRTL",
        win32con.WS_EX_LEFT: "WS_EX_LEFT",
        win32con.WS_EX_LEFTSCROLLBAR: "WS_EX_LEFTSCROLLBAR",
        win32con.WS_EX_LTRREADING: "WS_EX_LTRREADING",
        win32con.WS_EX_MDICHILD: "WS_EX_MDICHILD",
        win32con.WS_EX_NOACTIVATE: "WS_EX_NOACTIVATE",
        win32con.WS_EX_NOINHERITLAYOUT: "WS_EX_NOINHERITLAYOUT",
        win32con.WS_EX_NOPARENTNOTIFY: "WS_EX_NOPARENTNOTIFY",
        win32con.WS_EX_OVERLAPPEDWINDOW: "WS_EX_OVERLAPPEDWINDOW",
        win32con.WS_EX_PALETTEWINDOW: "WS_EX_PALETTEWINDOW",
        win32con.WS_EX_RIGHT: "WS_EX_RIGHT",
        win32con.WS_EX_RIGHTSCROLLBAR: "WS_EX_RIGHTSCROLLBAR",
        win32con.WS_EX_RTLREADING: "WS_EX_RTLREADING",
        win32con.WS_EX_STATICEDGE: "WS_EX_STATICEDGE",
        win32con.WS_EX_TOOLWINDOW: "WS_EX_TOOLWINDOW",
        win32con.WS_EX_TOPMOST: "WS_EX_TOPMOST",
        win32con.WS_EX_TRANSPARENT: "WS_EX_TRANSPARENT",
        win32con.WS_EX_WINDOWEDGE: "WS_EX_WINDOWEDGE"
    }
    for s in ex_style_dict:
        if ex_style & s:
            ex_styles.append(ex_style_dict[s])
    return ', '.join(ex_styles)

def get_control_info(hwnd):
    class_name = win32gui.GetClassName(hwnd)
    control_info = f"Control Class: {class_name}\n"
    
    if class_name in ["Button", "Edit", "Static", "ListBox", "ComboBox", "ScrollBar"]:
        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
        if class_name == "Button":
            button_styles = {
                win32con.BS_PUSHBUTTON: "Push Button",
                win32con.BS_DEFPUSHBUTTON: "Default Push Button",
                win32con.BS_CHECKBOX: "Checkbox",
                win32con.BS_AUTOCHECKBOX: "Auto Checkbox",
                win32con.BS_RADIOBUTTON: "Radio Button",
                win32con.BS_GROUPBOX: "Group Box",
            }
            for bs, name in button_styles.items():
                if style & bs:
                    control_info += f"Button Type: {name}\n"
                    break
        elif class_name == "Edit":
            if style & win32con.ES_MULTILINE:
                control_info += "Multiline Edit Control\n"
            else:
                control_info += "Single-line Edit Control\n"
        elif class_name == "Static":
            if style & win32con.SS_ICON:
                control_info += "Static Icon\n"
            elif style & win32con.SS_BITMAP:
                control_info += "Static Bitmap\n"
            else:
                control_info += "Static Text\n"
    
    return control_info

def get_menu_info(hwnd):
    menu_info = "Menu Information:\n"
    try:
        menu = win32gui.GetMenu(hwnd)
        if menu:
            menu_item_count = win32gui.GetMenuItemCount(menu)
            menu_info += f"Menu Items Count: {menu_item_count}\n"
            for i in range(menu_item_count):
                menu_state = win32gui.GetMenuState(menu, i, win32con.MF_BYPOSITION)
                menu_string = win32gui.GetMenuString(menu, i, win32con.MF_BYPOSITION)
                menu_info += f"  Item {i}: {menu_string} (State: {menu_state})\n"
        else:
            menu_info += "No menu found\n"
    except:
        menu_info += "Unable to retrieve menu information\n"
    return menu_info

def get_font_info(hwnd):
    font_info = "Font Information:\n"
    try:
        hdc = win32gui.GetDC(hwnd)
        font = win32gui.GetCurrentObject(hdc, win32con.OBJ_FONT)
        font_info += f"Font: {win32gui.GetObjectType(font)}\n"
        logfont = win32gui.GetObject(font)
        font_info += f"  Face Name: {logfont.lfFaceName}\n"
        font_info += f"  Height: {logfont.lfHeight}\n"
        font_info += f"  Weight: {logfont.lfWeight}\n"
        font_info += f"  Italic: {bool(logfont.lfItalic)}\n"
        font_info += f"  Underline: {bool(logfont.lfUnderline)}\n"
        font_info += f"  StrikeOut: {bool(logfont.lfStrikeOut)}\n"
        win32gui.ReleaseDC(hwnd, hdc)
    except:
        font_info += "Unable to retrieve font information\n"
    return font_info

def get_icon_info(hwnd):
    icon_info = "Icon Information:\n"
    try:
        icon = win32gui.GetClassLong(hwnd, win32con.GCL_HICON)
        if icon:
            icon_info += f"Icon Handle: {icon}\n"
            # You could potentially extract more information about the icon here
        else:
            icon_info += "No icon found\n"
    except:
        icon_info += "Unable to retrieve icon information\n"
    return icon_info

class WindowInfoViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Comprehensive Window and GUI Information Viewer")
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QHBoxLayout()
        central_widget.setLayout(layout)

        # Tree view for window hierarchy
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Window Title", "HWND"])
        self.tree.itemClicked.connect(self.on_item_clicked)
        layout.addWidget(self.tree, 1)

        right_layout = QVBoxLayout()
        layout.addLayout(right_layout, 2)

        # Search functionality
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        search_button = QPushButton("Search")
        search_button.clicked.connect(self.search_window)
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(search_button)
        right_layout.addLayout(search_layout)

        # Window info display
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.info_display = QTextEdit()
        self.info_display.setReadOnly(True)
        scroll_area.setWidget(self.info_display)
        right_layout.addWidget(scroll_area)

        # Click simulation controls
        click_layout = QHBoxLayout()
        self.x_input = QSpinBox()
        self.x_input.setRange(0, 5000)
        self.x_input.setPrefix("X: ")
        self.y_input = QSpinBox()
        self.y_input.setRange(0, 5000)
        self.y_input.setPrefix("Y: ")
        click_button = QPushButton("Click at X, Y")
        click_button.clicked.connect(self.click_at_position)
        click_layout.addWidget(self.x_input)
        click_layout.addWidget(self.y_input)
        click_layout.addWidget(click_button)
        right_layout.addLayout(click_layout)

        # Timer for updating the tree
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_tree)
        self.timer.start(1000)  # Update every 1 second

        self.update_tree()

    def update_tree(self):
        expanded_items = self.get_expanded_items()
        self.tree.clear()
        win32gui.EnumWindows(self.add_window_to_tree, 0)
        self.restore_expanded_items(expanded_items)

    def get_expanded_items(self):
        expanded = []
        def traverse(item):
            if item.isExpanded():
                expanded.append(item.text(1))
            for i in range(item.childCount()):
                traverse(item.child(i))
        for i in range(self.tree.topLevelItemCount()):
            traverse(self.tree.topLevelItem(i))
        return expanded

    def restore_expanded_items(self, expanded_items):
        def traverse(item):
            if item.text(1) in expanded_items:
                item.setExpanded(True)
            for i in range(item.childCount()):
                traverse(item.child(i))
        for i in range(self.tree.topLevelItemCount()):
            traverse(self.tree.topLevelItem(i))

    def add_window_to_tree(self, hwnd, lParam):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            item = QTreeWidgetItem([title, str(hwnd)])
            self.tree.addTopLevelItem(item)
            win32gui.EnumChildWindows(hwnd, self.add_child_to_tree, item)

    def add_child_to_tree(self, hwnd, item):
        title = win32gui.GetWindowText(hwnd)
        child_item = QTreeWidgetItem([title, str(hwnd)])
        item.addChild(child_item)
        win32gui.EnumChildWindows(hwnd, self.add_child_to_tree, child_item)
        return True

    def on_item_clicked(self, item, column):
        hwnd = int(item.text(1))
        self.show_window_info(hwnd)
        self.current_hwnd = hwnd

    def search_window(self):
        query = self.search_input.text().lower()
        for i in range(self.tree.topLevelItemCount()):
            item = self.tree.topLevelItem(i)
            self.search_item(item, query)

    def search_item(self, item, query):
        if query in item.text(0).lower() or query == item.text(1):
            self.tree.setCurrentItem(item)
            self.show_window_info(int(item.text(1)))
            return True
        for i in range(item.childCount()):
            if self.search_item(item.child(i), query):
                return True
        return False

    def show_window_info(self, hwnd):
        info = f"HWND: {hwnd}\n"
        info += f"Title: {win32gui.GetWindowText(hwnd)}\n"
        info += f"Class: {win32gui.GetClassName(hwnd)}\n"

        rect = win32gui.GetWindowRect(hwnd)
        info += f"Position: (left={rect[0]}, top={rect[1]}, right={rect[2]}, bottom={rect[3]})\n"
        info += f"Size: {rect[2]-rect[0]}x{rect[3]-rect[1]}\n"

        info += f"Window Styles: {get_window_style(hwnd)}\n"
        info += f"Extended Window Styles: {get_window_ex_style(hwnd)}\n"

        # GUI specific information
        info += "\n=== GUI Information ===\n"
        info += get_control_info(hwnd)
        info += get_menu_info(hwnd)
        info += get_font_info(hwnd)
        info += get_icon_info(hwnd)

        # Additional GUI elements
        info += "\nChild Windows:\n"
        child_windows = []
        win32gui.EnumChildWindows(hwnd, lambda hwnd, param: param.append(hwnd), child_windows)
        for child_hwnd in child_windows:
            info += f"  HWND: {child_hwnd}, Class: {win32gui.GetClassName(child_hwnd)}, Text: {win32gui.GetWindowText(child_hwnd)}\n"

        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            info += f"\n=== Process Information ===\n"
            info += f"Process Name: {process.name()}\n"
            info += f"Process ID: {pid}\n"
            info += f"Executable Path: {process.exe()}\n"
            info += f"Parent Process ID: {process.ppid()}\n"
            info += f"Creation Time: {process.create_time()}\n"
            info += f"CPU Usage: {process.cpu_percent()}%\n"
            memory_info = process.memory_info()
            info += f"Memory Usage: {memory_info.rss / (1024*1024):.2f} MB (RSS)\n"
            info += f"Virtual Memory: {memory_info.vms / (1024*1024):.2f} MB\n"
            info += f"Number of Threads: {process.num_threads()}\n"
            info += f"Priority: {process.nice()}\n"
            info += f"Status: {process.status()}\n"
            
            try:
                info += f"Command Line: {process.cmdline()}\n"
            except:
                info += "Command Line: Unable to retrieve\n"
            
            try:
                info += f"Current Working Directory: {process.cwd()}\n"
            except:
                info += "Current Working Directory: Unable to retrieve\n"
            
            try:
                info += f"Environment Variables: {process.environ()}\n"
            except:
                info += "Environment Variables: Unable to retrieve\n"
            
            try:
                info += f"Open Files: {process.open_files()}\n"
            except:
                info += "Open Files: Unable to retrieve\n"
            
            try:
                info += f"Connections: {process.connections()}\n"
            except:
                info += "Connections: Unable to retrieve\n"
            
        except Exception as e:
            info += f"Unable to retrieve process information: {str(e)}\n"

        self.info_display.setText(info)

    def click_at_position(self):
        if hasattr(self, 'current_hwnd'):
            x = self.x_input.value()
            y = self.y_input.value()

            # 해당 hwnd의 위치를 기준으로 x, y 좌표를 계산
            abs_x, abs_y, _, _ = win32gui.GetWindowRect(self.current_hwnd)
            
            lparam = win32api.MAKELONG(x + abs_x, y + abs_y)
            win32gui.SendMessage(self.current_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
            win32gui.SendMessage(self.current_hwnd, win32con.WM_LBUTTONUP, 0, lparam)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = WindowInfoViewer()
    viewer.show()
    sys.exit(app.exec_())
