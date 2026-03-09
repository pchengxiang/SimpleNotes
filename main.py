import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog, simpledialog
import json
import os
from dataclasses import dataclass, asdict
from typing import Dict, Any


@dataclass
class NoteData:
    id: int
    x: int
    y: int
    width: int
    height: int
    title: str
    text: str
    bg: str = "#fff9a5"
    line_spacing: int = 0  # 行距（像素），預設 0


class StickyNote:
    HANDLE_SIZE = 10

    def __init__(self, board, note_data: NoteData):
        self.board = board
        self.data = note_data
        self.drag_offset_x = 0
        self.drag_offset_y = 0
        self.resizing = False
        

        self.frame = tk.Frame(
            board.canvas,
            bg=self.data.bg,
            bd=1,
            relief="solid",
        )

        # 標題列
        self.title_bar = tk.Frame(self.frame, bg="#f0e68c")
        self.title_bar.pack(fill="x", side="top")

        self.title_label = tk.Button(
            self.title_bar,
            text=self.data.title,
            command=self.rename,
            bg="#f0e68c",
            anchor="w",
        )
        self.title_label.pack(side="left", padx=4)

        self.btn_delete = tk.Button(
            self.title_bar,
            text="✕",
            command=self.delete,
            bg="#f0e68c",
            relief="flat",
            padx=2,
            pady=0,
        )
        self.btn_delete.pack(side="right")

        self.btn_color = tk.Button(
            self.title_bar,
            text="🎨",
            command=self.change_color,
            bg="#f0e68c",
            relief="flat",
            padx=2,
            pady=0,
        )
        self.btn_color.pack(side="right")

        self.btn_spacing = tk.Button(
            self.title_bar,
            text="📏",
            command=self.change_line_spacing,
            bg="#f0e68c",
            relief="flat",
            padx=2,
            pady=0,
        )
        self.btn_spacing.pack(side="right")
        
        self.rename_entry = tk.Entry(
            self.title_bar,
            text=self.data.text,
            bg="#f0e68c",
            relief="flat"
        )        
        self.rename_entry.bind("<Return>", self.rename_entry_change)
        self.renaming = False

        # 文字區
        self.text_widget = tk.Text(
            self.frame,
            wrap="word",
            bg=self.data.bg,
            borderwidth=0,
            padx=4,
            pady=4,
            spacing2=self.data.line_spacing,  # 行距設定
        )
        self.text_widget.insert("1.0", self.data.text)
        self.text_widget.pack(fill="both", expand=True)

        # 右下角縮放把手
        self.resize_handle = tk.Frame(
            self.frame,
            bg="#d0c060",
            cursor="size_nw_se",
            width=self.HANDLE_SIZE,
            height=self.HANDLE_SIZE,
        )
        self.resize_handle.place(relx=1.0, rely=1.0, anchor="se")

        # 在 Canvas 上建立 window
        self.canvas_window = board.canvas.create_window(
            self.data.x,
            self.data.y,
            window=self.frame,
            anchor="nw",
            width=self.data.width,
            height=self.data.height,
        )

        # 綁定拖曳事件（移動）
        for widget in (self.title_bar, self.title_label):
            widget.bind("<Button-1>", self.on_drag_start)
            widget.bind("<B1-Motion>", self.on_drag_move)

        # 綁定縮放事件
        self.resize_handle.bind("<Button-1>", self.on_resize_start)
        self.resize_handle.bind("<B1-Motion>", self.on_resize_move)

        # 當點擊便條時，置於最上層
        self.frame.bind("<Button-1>", self.bring_to_front)
        self.text_widget.bind("<Button-1>", self.bring_to_front)
        self.title_bar.bind("<Button-1>", self.bring_to_front)

    def bring_to_front(self, event=None):
        self.board.canvas.tag_raise(self.canvas_window)

    def on_drag_start(self, event):
        self.bring_to_front()
        x, y = self.board.canvas.coords(self.canvas_window)
        self.drag_offset_x = event.x
        self.drag_offset_y = event.y
        self.start_x = x
        self.start_y = y

    def on_drag_move(self, event):
        dx = event.x - self.drag_offset_x
        dy = event.y - self.drag_offset_y
        new_x = self.start_x + dx
        new_y = self.start_y + dy

        # if new_x < 0:
        #     new_x = 0
        # if new_y < 0:
        #     new_y = 0
        # if new_x > self.board.canvas.winfo_width():
        #     new_x = self.board.canvas.winfo_width() - self.data.width
        # if new_y > self.board.canvas.winfo_height():
        #     new_y = self.board.canvas.winfo_height() - self.data.height
        self.board.canvas.coords(self.canvas_window, new_x, new_y)
        self.data.x = int(new_x)
        self.data.y = int(new_y)
        self.start_x = new_x
        self.start_y = new_y

    def on_resize_start(self, event):
        self.resizing = True
        bbox = self.board.canvas.bbox(self.canvas_window)
        if bbox:
            x1, y1, x2, y2 = bbox
            self.start_width = x2 - x1
            self.start_height = y2 - y1
        else:
            self.start_width = self.data.width
            self.start_height = self.data.height
        self.start_mouse_x = event.x_root
        self.start_mouse_y = event.y_root

    def on_resize_move(self, event):
        if not self.resizing:
            return
        dx = event.x_root - self.start_mouse_x
        dy = event.y_root - self.start_mouse_y
        new_width = max(120, self.start_width + dx)
        new_height = max(80, self.start_height + dy)
        self.board.canvas.itemconfig(self.canvas_window, width=new_width, height=new_height)
        self.data.width = int(new_width)
        self.data.height = int(new_height)

    def delete(self):
        if messagebox.askyesno("刪除便條", "確定要刪除這張便條嗎？"):
            self.board.remove_note(self.data.id)

    def rename(self):
        if self.renaming:
            self.renaming = False
            self.rename_entry.pack_forget()
            return
        self.rename_entry.pack(side="right")
        self.rename_entry.focus_set()
        self.rename_entry.select_range(0, "end")
        self.renaming = True

    def rename_entry_change(self, event=None):
        self.data.title = self.rename_entry.get()
        self.title_label.config(text=self.data.title)
        self.renaming = False
        self.rename_entry.pack_forget()

    def change_color(self):
        color = colorchooser.askcolor(color=self.data.bg, title="選擇便條顏色")
        if color and color[1]:
            self.data.bg = color[1]
            self.frame.config(bg=self.data.bg)
            self.text_widget.config(bg=self.data.bg)

    def change_line_spacing(self):
        """調整文字行距。"""
        current = self.data.line_spacing
        parent_win = self.board.root
        
        # 使用對話框輸入行距值（像素）
        value_str = simpledialog.askstring(
            "調整行距",
            f"請輸入行距（像素，目前：{current}）：\n（建議範圍：0-20，0 為預設行距）",
            initialvalue=str(current),
            parent=parent_win
        )
        
        if value_str is None:
            return
        
        try:
            spacing = int(value_str)
            # 限制在合理範圍
            spacing = max(0, min(50, spacing))
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數字。", parent=parent_win)
            return
        
        # 套用到 Text widget（spacing2 是全域設定，會自動套用到所有文字）
        self.data.line_spacing = spacing
        self.text_widget.tag_configure("line_spacing", spacing3=spacing)
        self.text_widget.tag_add("line_spacing", "1.0", "end")

    def get_state(self) -> Dict[str, Any]:
        self.data.text = self.text_widget.get("1.0", "end-1c")
        self.data.title = self.title_label.cget("text")
        return asdict(self.data)

    def destroy(self):
        self.board.canvas.delete(self.canvas_window)


class BoardApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("小紙條拼貼板 - Python")
        self.root.geometry("1000x700")

        self.current_id = 1
        self.notes: Dict[int, StickyNote] = {}
        # 目前開啟中的版面檔路徑（用於「儲存」時直接覆寫）
        self.current_board_path: str | None = None

        # 工具列與選單按鈕設定
        self.buttons_conf = [
            ["新增便條", self.add_note],
            ["儲存", self.save_board],          # 直接覆寫目前檔案（若有）
            ["另存新檔", self.save_board_as],   # 開啟對話框選擇新檔名
            ["開啟檔案", self.load_board],
            ["清空版面", self.clear_board],
            ["版面列表", self.open_board_list],
        ]

        self.file_menu_conf = [
            [["新增便條", self.add_note]],
            [["儲存", self.save_board],
             ["另存新檔", self.save_board_as],
             ["開啟檔案", self.load_board],
             ["版面列表", self.open_board_list]],
            [["離開", self.root.quit]],
        ]

        # 便條預設
        self.note_offset_x = 10
        self.note_offset_y = 10

        self.note_width = 220
        self.note_height = 160

        self.note_color = "#fff9a5"

        self.note_text = ""

        # 版面列表 / 分類相關
        self.notes_dir = os.path.join(os.path.dirname(__file__), "Notes")
        os.makedirs(self.notes_dir, exist_ok=True)
        self.board_list_window = None
        self.board_list_tree = None
        self.board_paths: Dict[str, str] = {}
        self.board_category_dirs: Dict[str, str | None] = {}

        self.create_widgets()
        self.create_menu()

    def create_widgets(self):
        # 上方工具列
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side="top", fill="x")

        for btn_text, btn_command in self.buttons_conf:
            btn = ttk.Button(toolbar, text=btn_text, command=btn_command)
            btn.pack(side="left", padx=4, pady=4)

        hint_label = ttk.Label(
            toolbar,
            text="提示：拖曳標題列可以移動，右下角可縮放，支援多張便條自由拼接排列。",
        )
        hint_label.pack(side="left", padx=12)

        # 主畫布（捲動區域）
        self.canvas = tk.Canvas(self.root, bg="#f8f8f8")
        self.canvas.pack(fill="both", expand=True, side="left")

        # 捲軸
        vbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        vbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=vbar.set, scrollregion=(0, 0, 2000, 2000))

        # 允許滑鼠滾輪捲動
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<Control-MouseWheel>", self.on_zoom)


    def create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
      
        for menu_list in self.file_menu_conf:
            for sub_menu_text, sub_menu_command in menu_list:
                file_menu.add_command(label=sub_menu_text, command=sub_menu_command)
            file_menu.add_separator()
        menubar.add_cascade(label="檔案", menu=file_menu)

        self.root.config(menu=menubar)

    def add_note(self, x: int = 40, y: int = 40):
        note_data = NoteData(
            id=self.current_id,
            x=x + self.current_id * self.note_offset_x,
            y=y + self.current_id * self.note_offset_y,
            width=self.note_width,
            height=self.note_height,
            title=f"便條 {self.current_id}",
            text=self.note_text,
        )
        note = StickyNote(self, note_data)
        self.notes[self.current_id] = note
        self.current_id += 1

    def remove_note(self, note_id: int):
        note = self.notes.pop(note_id, None)
        if note:
            note.destroy()

    def save_board(self):
        """一般儲存：若已開啟某檔案，直接覆寫；否則等同於另存新檔。"""
        if not self.notes:
            messagebox.showinfo("提醒", "目前沒有任何便條可儲存。")
            return

        if self.current_board_path:
            # 直接覆寫目前開啟中的檔案
            self.save_board_to_path(self.current_board_path, show_message=True)
        else:
            # 第一次儲存，視為另存新檔
            self.save_board_as()

    def save_board_as(self):
        """另存新檔：總是跳出對話框選擇存檔路徑，並更新目前檔案路徑。"""
        if not self.notes:
            messagebox.showinfo("提醒", "目前沒有任何便條可儲存。")
            return

        file_path = filedialog.asksaveasfilename(
            title="另存新檔",
            defaultextension=".json",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not file_path:
            return

        # 更新目前開啟中的檔案路徑
        self.current_board_path = file_path
        self.save_board_to_path(file_path, show_message=True)

    def save_board_to_path(self, file_path: str, show_message: bool = False):
        if not self.notes:
            messagebox.showinfo("提醒", "目前沒有任何便條可儲存。")
            return

        data = {
            "notes": [note.get_state() for note in self.notes.values()],
            "current_id": self.current_id,
        }
        try:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            if show_message:
                messagebox.showinfo("完成", "版面已成功儲存。")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存失敗：{e}")

    def clear_board(self):
        for note in list(self.notes.values()):
            note.destroy()
        self.notes.clear()
        self.current_id = 1
        self.file_path = None

    def load_board(self):
        file_path = filedialog.askopenfilename(
            title="載入版面",
            filetypes=[("JSON 檔案", "*.json"), ("所有檔案", "*.*")],
        )
        if not file_path:
            return

        self.load_board_from_path(file_path, show_message=True)

    def load_board_from_path(self, file_path: str, show_message: bool = False):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取失敗：{e}")
            return
        self.root.title(os.path.basename(file_path))
        # 清空現有便條
        for note in list(self.notes.values()):
            note.destroy()
        self.notes.clear()

        # 還原
        notes_data = data.get("notes", [])
        self.current_id = data.get("current_id", 1)

        for item in notes_data:
            try:
                note_data = NoteData(**item)
                note = StickyNote(self, note_data)
                self.notes[note_data.id] = note
            except TypeError:
                continue

        # 記住目前開啟的檔案，之後「儲存」會覆寫此檔
        self.current_board_path = file_path

        if show_message:
            messagebox.showinfo("完成", "版面已載入。")

    # === 版面列表 / 分類 ===
    def open_board_list(self):
        """打開版面列表視窗，可以用資料夾分類管理 .json 版面檔。"""
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()
            self.refresh_board_list()
            return

        win = tk.Toplevel(self.root)
        win.title("版面列表")
        win.geometry("500x400")
        self.board_list_window = win

        # 上方按鈕列
        top_bar = ttk.Frame(win)
        top_bar.pack(side="top", fill="x", padx=6, pady=4)

        btn_add_folder = ttk.Button(top_bar, text="新增資料夾", command=self.board_list_add_folder)
        btn_add_folder.pack(side="left", padx=2)

        btn_save_current = ttk.Button(top_bar, text="將目前版面加入列表", command=self.board_list_save_current)
        btn_save_current.pack(side="left", padx=2)

        btn_open = ttk.Button(top_bar, text="開啟選取版面", command=self.board_list_open_selected)
        btn_open.pack(side="left", padx=2)

        btn_move = ttk.Button(top_bar, text="移至其他資料夾", command=self.board_list_move_workspace)
        btn_move.pack(side="left", padx=2)

        btn_delete = ttk.Button(top_bar, text="刪除選取項目", command=self.board_list_delete_selected)
        btn_delete.pack(side="left", padx=2)

        # Treeview 顯示 資料夾 / 版面
        tree = ttk.Treeview(win)
        tree.heading("#0", text="資料夾 / 版面", anchor="w")
        tree.pack(side="top", fill="both", expand=True, padx=6, pady=(0, 6))
        tree.bind("<Double-1>", self.board_list_on_double_click)

        self.board_list_tree = tree
        self.refresh_board_list()

        def on_close():
            # 關閉視窗並重置引用，避免無法再次正確開啟
            if self.board_list_window is not None:
                self.board_list_window.destroy()
                self.board_list_window = None

        win.protocol("WM_DELETE_WINDOW", on_close)

    def refresh_board_list(self):
        """掃描 Notes 資料夾，重建樹狀結構。"""
        if self.board_list_tree is None:
            return

        # 清空
        for item in self.board_list_tree.get_children():
            self.board_list_tree.delete(item)
        self.board_paths.clear()
        self.board_category_dirs.clear()

        os.makedirs(self.notes_dir, exist_ok=True)

        # 未分類節點（Notes 根目錄底下的 .json）
        uncat_id = self.board_list_tree.insert("", "end", text="(未分類)", open=True)
        self.board_category_dirs[uncat_id] = None  # None 代表 Notes 根目錄

        for fname in sorted(os.listdir(self.notes_dir)):
            full_path = os.path.join(self.notes_dir, fname)
            if os.path.isdir(full_path):
                continue
            if not fname.lower().endswith(".json"):
                continue
            ws_name = os.path.splitext(fname)[0]
            item_id = self.board_list_tree.insert(uncat_id, "end", text=ws_name)
            self.board_paths[item_id] = full_path

        # 子資料夾作為分類
        for entry in sorted(os.listdir(self.notes_dir)):
            cat_path = os.path.join(self.notes_dir, entry)
            if not os.path.isdir(cat_path):
                continue

            cat_id = self.board_list_tree.insert("", "end", text=entry, open=True)
            self.board_category_dirs[cat_id] = cat_path

            for fname in sorted(os.listdir(cat_path)):
                if not fname.lower().endswith(".json"):
                    continue
                ws_name = os.path.splitext(fname)[0]
                full_path = os.path.join(cat_path, fname)
                item_id = self.board_list_tree.insert(cat_id, "end", text=ws_name)
                self.board_paths[item_id] = full_path

    def board_list_add_folder(self):
        """在 Notes 底下新增一個資料夾分類。"""
        parent_win = self.board_list_window or self.root
        name = simpledialog.askstring("新增資料夾", "請輸入資料夾名稱：", parent=parent_win)
        if not name:
            return
        name = name.strip()
        if not name:
            return

        dir_path = os.path.join(self.notes_dir, name)
        try:
            os.makedirs(dir_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("錯誤", f"建立資料夾失敗：{e}", parent=parent_win)
            return

        self.refresh_board_list()
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()

    def _resolve_selected_category_dir(self):
        """從目前選取的 Tree item 推出對應的資料夾路徑。"""
        if self.board_list_tree is None:
            return None
        selection = self.board_list_tree.selection()
        if not selection:
            return None
        sel_item = selection[0]

        # 如果選到的是版面（檔案），就往上找分類節點
        if sel_item in self.board_paths:
            cat_item = self.board_list_tree.parent(sel_item)
        else:
            cat_item = sel_item

        if cat_item not in self.board_category_dirs:
            return None

        base_dir = self.board_category_dirs[cat_item]
        if base_dir is None:
            return self.notes_dir
        return base_dir

    def board_list_save_current(self):
        """將目前畫面存成一個 .json 加入列表（依選取的資料夾分類）。"""
        parent_win = self.board_list_window or self.root
        if not self.notes:
            messagebox.showinfo("提醒", "目前沒有任何便條可儲存。", parent=parent_win)
            return

        target_dir = self._resolve_selected_category_dir()
        if not target_dir:
            messagebox.showinfo("提示", "請先在列表中選擇一個資料夾，或選擇「(未分類)」。", parent=parent_win)
            return

        name = simpledialog.askstring("儲存版面", "請輸入此版面的名稱：", parent=parent_win)
        if not name:
            return
        name = name.strip()
        if not name:
            return

        os.makedirs(target_dir, exist_ok=True)
        file_path = os.path.join(target_dir, name + ".json")

        # 若檔案已存在，詢問是否覆寫
        if os.path.exists(file_path):
            if not messagebox.askyesno("覆寫確認", "同名檔案已存在，確定要覆寫嗎？", parent=parent_win):
                return

        self.save_board_to_path(file_path, show_message=True)
        self.refresh_board_list()
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()

    def board_list_open_selected(self):
        """從列表中載入選取的版面。"""
        parent_win = self.board_list_window or self.root
        if self.board_list_tree is None:
            return
        selection = self.board_list_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "請先在列表中選擇一個版面。", parent=parent_win)
            return
        item_id = selection[0]

        path = self.board_paths.get(item_id)
        if not path:
            messagebox.showinfo("提示", "請選擇一個實際的版面檔（不是資料夾）。", parent=parent_win)
            return

        self.load_board_from_path(path, show_message=True)
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()

    def board_list_delete_selected(self):
        """刪除選取的版面檔，或刪除空資料夾。"""
        parent_win = self.board_list_window or self.root
        if self.board_list_tree is None:
            return
        selection = self.board_list_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "請先選擇要刪除的項目。", parent=parent_win)
            return
        item_id = selection[0]

        # 先看是不是檔案
        path = self.board_paths.get(item_id)
        if path:
            if not messagebox.askyesno("刪除確認", f"確定要刪除版面檔：\n{os.path.basename(path)}？", parent=parent_win):
                return
            try:
                os.remove(path)
            except Exception as e:
                messagebox.showerror("錯誤", f"刪除檔案失敗：{e}", parent=parent_win)
                return
            self.refresh_board_list()
            if self.board_list_window is not None and self.board_list_window.winfo_exists():
                self.board_list_window.lift()
                self.board_list_window.focus_force()
            return

        # 否則當成資料夾處理（不能刪除未分類）
        if item_id not in self.board_category_dirs:
            return
        dir_path = self.board_category_dirs[item_id]
        if dir_path is None:
            messagebox.showinfo("提示", "「(未分類)」是根節點，不能刪除。", parent=parent_win)
            return

        if os.listdir(dir_path):
            messagebox.showinfo("提示", "此資料夾非空，請先刪除底下所有版面檔再嘗試刪除。", parent=parent_win)
            return

        if not messagebox.askyesno("刪除確認", f"確定要刪除資料夾：\n{os.path.basename(dir_path)}？", parent=parent_win):
            return

        try:
            os.rmdir(dir_path)
        except Exception as e:
            messagebox.showerror("錯誤", f"刪除資料夾失敗：{e}", parent=parent_win)
            return

        self.refresh_board_list()
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()

    def board_list_move_workspace(self):
        """將選取的版面檔移動到其他資料夾（或未分類）。

        使用方式：在列表中「同時選取」一個工作區與一個目標資料夾（可用 Ctrl 多選），再按此按鈕。
        """
        parent_win = self.board_list_window or self.root
        if self.board_list_tree is None:
            return

        selection = self.board_list_tree.selection()
        if not selection:
            messagebox.showinfo(
                "提示",
                "請在列表中同時選擇一個「要移動的工作區」以及一個「目標資料夾」（可用 Ctrl 多選）。",
                parent=parent_win,
            )
            return

        work_items = [i for i in selection if i in self.board_paths]
        cat_items = [i for i in selection if i in self.board_category_dirs]

        if len(work_items) != 1 or len(cat_items) != 1:
            messagebox.showinfo(
                "提示",
                "請確保「恰好選取一個工作區檔案」以及「恰好選取一個資料夾」後再執行移動。\n"
                "（可在列表中按住 Ctrl 來多選）",
                parent=parent_win,
            )
            return

        work_item = work_items[0]
        cat_item = cat_items[0]

        src_path = self.board_paths.get(work_item)
        if not src_path:
            messagebox.showinfo("提示", "無法取得要移動的工作區檔案。", parent=parent_win)
            return

        base_dir = self.board_category_dirs.get(cat_item)
        if base_dir is None:
            target_dir = self.notes_dir  # (未分類)
        else:
            target_dir = base_dir

        os.makedirs(target_dir, exist_ok=True)
        dest_path = os.path.join(target_dir, os.path.basename(src_path))

        if os.path.abspath(dest_path) == os.path.abspath(src_path):
            messagebox.showinfo("提示", "目標資料夾與目前位置相同，沒有需要移動。", parent=parent_win)
            return

        if os.path.exists(dest_path):
            if not messagebox.askyesno(
                "覆寫確認",
                f"目標資料夾中已存在同名檔案：\n{os.path.basename(dest_path)}\n\n確定要覆寫嗎？",
                parent=parent_win,
            ):
                return

        try:
            os.replace(src_path, dest_path)
        except Exception as e:
            messagebox.showerror("錯誤", f"移動檔案失敗：{e}", parent=parent_win)
            return

        self.refresh_board_list()
        if self.board_list_window is not None and self.board_list_window.winfo_exists():
            self.board_list_window.lift()
            self.board_list_window.focus_force()

    def board_list_on_double_click(self, event):
        """在列表中雙擊時，若是版面檔就直接開啟。"""
        if self.board_list_tree is None:
            return
        item_id = self.board_list_tree.focus()
        if not item_id:
            return
        path = self.board_paths.get(item_id)
        if not path:
            return
        self.load_board_from_path(path, show_message=True)

    def on_mouse_wheel(self, event):
        # Windows 滾輪方向
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def zoom(self, factor):
        # 以 (0, 0) 為中心縮放，也可以改成視窗中心
        self.canvas.scale("all", 0, 0, factor, factor)
        # 記得同步調整 scrollregion，否則捲軸範圍會錯
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_zoom(self, event):
        if event.delta > 0:
            factor = 1.1
        else:
            factor = 0.9
        self.zoom(factor)

def main():
    root = tk.Tk()
    app = BoardApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

