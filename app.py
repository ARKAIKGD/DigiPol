import os
import re
import io
import sys
import threading
import tkinter as tk
from csv import writer
from datetime import datetime
from queue import Empty, Queue
from tkinter import filedialog, messagebox, simpledialog, ttk

from PIL import Image, ImageDraw, ImageGrab, ImageTk


SAVE_DIR = os.path.join(os.path.expanduser("~"), "Pictures", "StudentSnips")
LOG_PATH = os.path.join(SAVE_DIR, "capture_log.csv")
VERSION_FILE = os.path.join(os.path.dirname(__file__), "version.txt")
ICON_REL_PATH = os.path.join("assets", "studentsnip.ico")
MAX_GIF_BYTES = 50 * 1024 * 1024


def get_app_version():
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as version_file:
            value = version_file.read().strip()
            return value if value else "dev"
    except OSError:
        return "dev"


def get_build_timestamp():
    if getattr(sys, "frozen", False):
        candidate_paths = [sys.executable]
    else:
        candidate_paths = [VERSION_FILE, __file__]

    for candidate_path in candidate_paths:
        try:
            modified_at = datetime.fromtimestamp(os.path.getmtime(candidate_path))
            return modified_at.strftime("%Y-%m-%d %H:%M")
        except OSError:
            continue

    return "unknown"


class HoverToolTip:
    def __init__(self, widget, text, delay_ms=3000):
        self.widget = widget
        self.text = text
        self.delay_ms = delay_ms
        self.after_id = None
        self.tip_window = None

        self.widget.bind("<Enter>", self._schedule, add="+")
        self.widget.bind("<Leave>", self._hide, add="+")
        self.widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _event=None):
        self._cancel_schedule()
        self.after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel_schedule(self):
        if self.after_id is not None:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def _show(self):
        if self.tip_window is not None:
            return

        root_x = self.widget.winfo_rootx()
        root_y = self.widget.winfo_rooty()
        x = root_x + 12
        y = root_y + self.widget.winfo_height() + 8

        tip_window = tk.Toplevel(self.widget)
        tip_window.wm_overrideredirect(True)
        tip_window.attributes("-topmost", True)
        tip_window.wm_geometry(f"+{x}+{y}")

        label = tk.Label(
            tip_window,
            text=self.text,
            justify="left",
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 8),
            padx=6,
            pady=3,
            wraplength=260,
        )
        label.pack()

        self.tip_window = tip_window

    def _hide(self, _event=None):
        self._cancel_schedule()
        if self.tip_window is not None:
            self.tip_window.destroy()
            self.tip_window = None


class ProgressFileWriter:
    def __init__(self, file_path, progress_callback=None, min_update_bytes=262144, cancel_event=None):
        self.file_path = file_path
        self.progress_callback = progress_callback
        self.min_update_bytes = max(1, int(min_update_bytes))
        self.cancel_event = cancel_event
        self._file = open(file_path, "wb")
        self.bytes_written = 0
        self._last_reported = 0
        self.name = file_path

    def write(self, data):
        if self.cancel_event is not None and self.cancel_event.is_set():
            raise RuntimeError("CANCELED_BY_USER")
        written = self._file.write(data)
        self.bytes_written += written
        if self.progress_callback is not None:
            if (self.bytes_written - self._last_reported) >= self.min_update_bytes:
                self._last_reported = self.bytes_written
                self.progress_callback(self.bytes_written)
        return written

    def flush(self):
        self._file.flush()

    def tell(self):
        return self._file.tell()

    def seek(self, offset, whence=0):
        return self._file.seek(offset, whence)

    def close(self):
        if not self._file.closed:
            self._file.close()

    def writable(self):
        return True

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


class SnippingTool:
    def __init__(self):
        self.app_version = get_app_version()
        self.build_timestamp = get_build_timestamp()
        self._tooltips = []
        self.icon_path = self._resolve_icon_path()
        self._configure_windows_app_id()
        self.root = tk.Tk()
        self.root.title(
            f"Student Screenshot Tool v{self.app_version} | Build {self.build_timestamp}"
        )
        self.root.resizable(False, False)
        self.root.bind_all("<Control-n>", self._on_shortcut_start_snip)
        self.root.bind_all("<Control-o>", self._on_shortcut_open_folder)
        self._apply_window_icon(self.root)

        self.status_var = tk.StringVar(value="Ready")
        self.start_x = 0
        self.start_y = 0
        self.end_x = 0
        self.end_y = 0
        self.start_canvas_x = 0
        self.start_canvas_y = 0
        self.end_canvas_x = 0
        self.end_canvas_y = 0
        self.preview_state = {}
        self.progress_frame_window = None
        self.progress_frames = []
        self.gif_preview_state = {}
        self.video_capture_running = False
        self.video_capture_after_id = None
        self.video_capture_interval_ms = 500
        self.short_video_mode_armed = False
        self.preferred_preview_speed_ms = 500
        self.save_progress_state = {}
        self.progress_frame_drag_state = {}

        self._build_main_ui()
        self._fit_window_to_content()

    def _resolve_icon_path(self):
        if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, ICON_REL_PATH)
        return os.path.join(os.path.dirname(__file__), ICON_REL_PATH)

    def _configure_windows_app_id(self):
        if os.name != "nt":
            return
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DigiPol.StudentSnip")
        except Exception:
            pass

    def _apply_window_icon(self, window):
        if not self.icon_path or not os.path.exists(self.icon_path):
            return
        try:
            window.iconbitmap(self.icon_path)
        except Exception:
            pass

    def _add_tooltip(self, widget, text):
        self._tooltips.append(HoverToolTip(widget, text, delay_ms=3000))

    def _fit_window_to_content(self):
        self.root.update_idletasks()
        required_width = self.root.winfo_reqwidth()
        required_height = self.root.winfo_reqheight()
        final_width = required_width + 2
        final_height = max(required_height, 180)
        self.root.geometry(f"{final_width}x{final_height}")

    def _build_main_ui(self):
        container = tk.Frame(self.root, padx=0, pady=4)
        container.pack(fill="both", expand=True)

        title = tk.Label(
            container,
            text="Screenshot Snipping Tool",
            font=("Segoe UI", 10, "bold"),
        )
        title.pack(pady=(0, 1))

        build_info = tk.Label(
            container,
            text=f"v{self.app_version} • Build {self.build_timestamp}",
            justify="center",
            font=("Segoe UI", 7),
            fg="#4a4a4a",
        )
        build_info.pack(pady=(0, 2))

        shortcuts = tk.Label(
            container,
            text="Ctrl+N Snip  •  Ctrl+O Folder",
            wraplength=240,
            justify="center",
            font=("Segoe UI", 7),
            fg="#444444",
        )
        shortcuts.pack(pady=(0, 3))

        actions = tk.Frame(container)
        actions.pack(pady=(0, 3), padx=0)

        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)

        start_btn = tk.Button(
            actions,
            text="Start Snip",
            width=13,
            command=self.start_snip,
            font=("Segoe UI", 7),
        )
        start_btn.grid(row=0, column=0, padx=1, pady=1)
        self._add_tooltip(start_btn, "Start manual snip selection (shortcut: Ctrl+N).")

        folder_btn = tk.Button(
            actions,
            text="Open Folder",
            width=13,
            command=self.open_save_folder,
            font=("Segoe UI", 7),
        )
        folder_btn.grid(row=0, column=1, padx=1, pady=1)
        self._add_tooltip(folder_btn, "Open the StudentSnips save folder (shortcut: Ctrl+O).")

        frame_btn = tk.Button(
            actions,
            text="Camera Frame",
            width=13,
            command=self.open_progress_frame,
            font=("Segoe UI", 7),
        )
        frame_btn.grid(row=1, column=0, padx=1, pady=1)
        self._add_tooltip(frame_btn, "Create or focus the movable camera frame over your work area.")

        capture_frame_btn = tk.Button(
            actions,
            text="Capture Frame",
            width=13,
            command=self.capture_progress_frame,
            font=("Segoe UI", 7),
        )
        capture_frame_btn.grid(row=1, column=1, padx=1, pady=1)
        self._add_tooltip(capture_frame_btn, "Capture one frame from the camera frame region.")

        short_video_btn = tk.Button(
            actions,
            text="Short Video",
            width=13,
            command=self.begin_short_video_mode,
            font=("Segoe UI", 7),
        )
        short_video_btn.grid(row=2, column=0, padx=1, pady=1)
        self._add_tooltip(short_video_btn, "Set FPS for timed capture. Then click Capture Frame once to start recording.")

        stop_capture_btn = tk.Button(
            actions,
            text="Stop Capture",
            width=13,
            command=self.stop_short_video_capture,
            font=("Segoe UI", 7),
        )
        stop_capture_btn.grid(row=2, column=1, padx=1, pady=1)
        self._add_tooltip(stop_capture_btn, "Stop timed capture and open preview if enough frames were recorded.")

        clear_frames_btn = tk.Button(
            actions,
            text="Clear Frames",
            width=13,
            command=self.clear_progress_frames,
            font=("Segoe UI", 7),
        )
        clear_frames_btn.grid(row=3, column=0, padx=1, pady=1)
        self._add_tooltip(clear_frames_btn, "Remove all currently captured progress frames.")

        export_gif_btn = tk.Button(
            actions,
            text="Export Progress",
            width=13,
            command=self.export_progress_gif,
            font=("Segoe UI", 7),
        )
        export_gif_btn.grid(row=3, column=1, padx=1, pady=1)
        self._add_tooltip(export_gif_btn, "Open preview to choose frames and save as GIF, MP4, or WebP.")

        status = tk.Label(container, textvariable=self.status_var, fg="#333333", font=("Segoe UI", 7))
        status.pack(pady=(3, 0))

    def start_snip(self):
        self.status_var.set("Drag to select an area...")
        self.root.withdraw()
        self.root.after(180, self._open_overlay)

    def _on_shortcut_start_snip(self, _event=None):
        self.start_snip()
        return "break"

    def _on_shortcut_open_folder(self, _event=None):
        self.open_save_folder()
        return "break"

    def _open_overlay(self):
        self.overlay = tk.Toplevel()
        self.overlay.attributes("-fullscreen", True)
        self.overlay.attributes("-alpha", 0.25)
        self.overlay.attributes("-topmost", True)
        self.overlay.configure(bg="black")
        self.overlay.title("Select region")
        self._apply_window_icon(self.overlay)

        self.canvas = tk.Canvas(self.overlay, cursor="cross", bg="gray11", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.rect_id = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.overlay.bind("<Escape>", self.cancel_snip)

    def on_button_press(self, event):
        self.start_x, self.start_y = event.x_root, event.y_root
        self.end_x, self.end_y = event.x_root, event.y_root
        self.start_canvas_x, self.start_canvas_y = event.x, event.y
        self.end_canvas_x, self.end_canvas_y = event.x, event.y
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        self.rect_id = self.canvas.create_rectangle(
            self.start_canvas_x,
            self.start_canvas_y,
            self.end_canvas_x,
            self.end_canvas_y,
            outline="red",
            width=2,
        )

    def on_drag(self, event):
        self.end_x, self.end_y = event.x_root, event.y_root
        self.end_canvas_x, self.end_canvas_y = event.x, event.y
        if self.rect_id:
            self.canvas.coords(
                self.rect_id,
                self.start_canvas_x,
                self.start_canvas_y,
                self.end_canvas_x,
                self.end_canvas_y,
            )

    def on_button_release(self, _event):
        x1, y1 = min(self.start_x, self.end_x), min(self.start_y, self.end_y)
        x2, y2 = max(self.start_x, self.end_x), max(self.start_y, self.end_y)

        self.overlay.destroy()

        if abs(x2 - x1) < 2 or abs(y2 - y1) < 2:
            self.status_var.set("Selection too small. Try again.")
            self.root.deiconify()
            self.root.lift()
            return

        try:
            image = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
            self._show_preview(image)
        except Exception as exc:
            messagebox.showerror("Capture Error", f"Failed to capture screenshot:\n{exc}")
            self.status_var.set("Capture failed")
            self.root.deiconify()
            self.root.lift()

    def _show_preview(self, image):
        self.root.deiconify()
        self.root.lift()

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Preview Screenshot")
        preview_window.resizable(False, False)
        preview_window.transient(self.root)
        preview_window.grab_set()
        self._apply_window_icon(preview_window)

        preview_container = tk.Frame(preview_window, padx=12, pady=12)
        preview_container.pack(fill="both", expand=True)

        image_width, image_height = image.size
        max_width, max_height = 900, 600
        scale = min(max_width / image_width, max_height / image_height, 1.0)
        display_width = max(int(image_width * scale), 1)
        display_height = max(int(image_height * scale), 1)

        tools_row = tk.Frame(preview_container)
        tools_row.pack(pady=(0, 8))

        tool_var = tk.StringVar(value="draw")
        tk.Radiobutton(tools_row, text="Draw", value="draw", variable=tool_var).pack(side="left", padx=4)
        tk.Radiobutton(tools_row, text="Rectangle", value="rect", variable=tool_var).pack(side="left", padx=4)
        tk.Radiobutton(tools_row, text="Text", value="text", variable=tool_var).pack(side="left", padx=4)

        text_var = tk.StringVar(value="Note")
        text_entry = tk.Entry(tools_row, textvariable=text_var, width=18)
        text_entry.pack(side="left", padx=(8, 4))

        canvas = tk.Canvas(
            preview_container,
            width=display_width,
            height=display_height,
            bg="white",
            highlightthickness=1,
            highlightbackground="#a0a0a0",
        )
        canvas.pack(pady=(0, 10))

        self.preview_state = {
            "window": preview_window,
            "canvas": canvas,
            "image": image.copy(),
            "base_image": image.copy(),
            "scale": scale,
            "tool_var": tool_var,
            "text_var": text_var,
            "last_xy": None,
            "rect_start": None,
            "temp_rect": None,
            "undo_stack": [],
        }

        self._refresh_preview_canvas()

        canvas.bind("<ButtonPress-1>", self._preview_on_press)
        canvas.bind("<B1-Motion>", self._preview_on_drag)
        canvas.bind("<ButtonRelease-1>", self._preview_on_release)

        info_label = tk.Label(
            preview_container,
            text="Annotate if needed, then Save. Files are saved in step order to Pictures\\StudentSnips.",
            justify="center",
            wraplength=880,
        )
        info_label.pack(pady=(0, 10))

        button_row = tk.Frame(preview_container)
        button_row.pack()

        undo_btn = tk.Button(
            button_row,
            text="Undo",
            width=10,
            command=self._preview_undo,
        )
        undo_btn.pack(side="left", padx=4)
        self._add_tooltip(undo_btn, "Undo the most recent annotation change.")

        clear_btn = tk.Button(
            button_row,
            text="Clear",
            width=10,
            command=self._preview_clear,
        )
        clear_btn.pack(side="left", padx=4)
        self._add_tooltip(clear_btn, "Reset annotations and return to the original captured image.")

        save_btn = tk.Button(
            button_row,
            text="Save",
            width=14,
            command=lambda: self._save_image(preview_window),
        )
        save_btn.pack(side="left", padx=4)
        self._add_tooltip(save_btn, "Save this screenshot as the next step image.")

        cancel_btn = tk.Button(
            button_row,
            text="Cancel",
            width=14,
            command=lambda: self._cancel_preview(preview_window),
        )
        cancel_btn.pack(side="left", padx=4)
        self._add_tooltip(cancel_btn, "Close preview without saving this screenshot.")

        preview_window.protocol("WM_DELETE_WINDOW", lambda: self._cancel_preview(preview_window))

    def _preview_push_undo(self):
        if not self.preview_state:
            return
        undo_stack = self.preview_state["undo_stack"]
        undo_stack.append(self.preview_state["image"].copy())
        if len(undo_stack) > 20:
            undo_stack.pop(0)

    def _preview_to_image_coords(self, canvas_x, canvas_y):
        scale = self.preview_state.get("scale", 1.0)
        image = self.preview_state["image"]
        max_x = image.width - 1
        max_y = image.height - 1
        image_x = int(canvas_x / scale)
        image_y = int(canvas_y / scale)
        image_x = min(max(image_x, 0), max_x)
        image_y = min(max(image_y, 0), max_y)
        return image_x, image_y

    def _refresh_preview_canvas(self):
        if not self.preview_state:
            return

        canvas = self.preview_state["canvas"]
        image = self.preview_state["image"]
        scale = self.preview_state["scale"]
        display_size = (max(int(image.width * scale), 1), max(int(image.height * scale), 1))
        display_image = image.resize(display_size)
        preview_photo = ImageTk.PhotoImage(display_image)
        self.preview_state["preview_photo"] = preview_photo
        canvas.delete("all")
        canvas.create_image(0, 0, image=preview_photo, anchor="nw")

    def _preview_on_press(self, event):
        if not self.preview_state:
            return

        tool = self.preview_state["tool_var"].get()
        canvas = self.preview_state["canvas"]

        if tool == "draw":
            self._preview_push_undo()
            self.preview_state["last_xy"] = (event.x, event.y)
        elif tool == "rect":
            self._preview_push_undo()
            self.preview_state["rect_start"] = (event.x, event.y)
            self.preview_state["temp_rect"] = canvas.create_rectangle(
                event.x,
                event.y,
                event.x,
                event.y,
                outline="red",
                width=2,
            )
        elif tool == "text":
            self._preview_push_undo()
            text_value = self.preview_state["text_var"].get().strip() or "Note"
            image_x, image_y = self._preview_to_image_coords(event.x, event.y)
            draw = ImageDraw.Draw(self.preview_state["image"])
            draw.text((image_x, image_y), text_value, fill="red")
            self._refresh_preview_canvas()

    def _preview_on_drag(self, event):
        if not self.preview_state:
            return

        tool = self.preview_state["tool_var"].get()
        canvas = self.preview_state["canvas"]

        if tool == "draw":
            last_xy = self.preview_state.get("last_xy")
            if last_xy is None:
                return
            x1, y1 = self._preview_to_image_coords(last_xy[0], last_xy[1])
            x2, y2 = self._preview_to_image_coords(event.x, event.y)
            draw = ImageDraw.Draw(self.preview_state["image"])
            draw.line((x1, y1, x2, y2), fill="red", width=4)
            self.preview_state["last_xy"] = (event.x, event.y)
            self._refresh_preview_canvas()
        elif tool == "rect":
            temp_rect = self.preview_state.get("temp_rect")
            rect_start = self.preview_state.get("rect_start")
            if temp_rect and rect_start:
                canvas.coords(temp_rect, rect_start[0], rect_start[1], event.x, event.y)

    def _preview_on_release(self, event):
        if not self.preview_state:
            return

        tool = self.preview_state["tool_var"].get()
        canvas = self.preview_state["canvas"]

        if tool == "draw":
            self.preview_state["last_xy"] = None
        elif tool == "rect":
            rect_start = self.preview_state.get("rect_start")
            temp_rect = self.preview_state.get("temp_rect")
            if rect_start:
                x1, y1 = self._preview_to_image_coords(rect_start[0], rect_start[1])
                x2, y2 = self._preview_to_image_coords(event.x, event.y)
                draw = ImageDraw.Draw(self.preview_state["image"])
                draw.rectangle((x1, y1, x2, y2), outline="red", width=4)
            if temp_rect:
                canvas.delete(temp_rect)
            self.preview_state["rect_start"] = None
            self.preview_state["temp_rect"] = None
            self._refresh_preview_canvas()

    def _preview_undo(self):
        if not self.preview_state:
            return
        undo_stack = self.preview_state["undo_stack"]
        if not undo_stack:
            return
        self.preview_state["image"] = undo_stack.pop()
        self._refresh_preview_canvas()

    def _preview_clear(self):
        if not self.preview_state:
            return
        self._preview_push_undo()
        if "base_image" not in self.preview_state:
            self.preview_state["base_image"] = self.preview_state["image"].copy()
        self.preview_state["image"] = self.preview_state["base_image"].copy()
        self._refresh_preview_canvas()

    def _next_step_number(self):
        os.makedirs(SAVE_DIR, exist_ok=True)
        max_step = 0
        for file_name in os.listdir(SAVE_DIR):
            match = re.match(r"step_(\d{3,})_", file_name)
            if match:
                max_step = max(max_step, int(match.group(1)))
        return max_step + 1

    def _write_capture_log(self, step_number, file_name):
        file_exists = os.path.exists(LOG_PATH)
        with open(LOG_PATH, "a", newline="", encoding="utf-8") as log_file:
            csv_writer = writer(log_file)
            if not file_exists:
                csv_writer.writerow(["step", "timestamp", "file_name", "full_path"])
            csv_writer.writerow(
                [
                    step_number,
                    datetime.now().isoformat(timespec="seconds"),
                    file_name,
                    os.path.join(SAVE_DIR, file_name),
                ]
            )

    def _save_image(self, preview_window=None):
        image = self.preview_state.get("image") if self.preview_state else None
        if image is None:
            self.status_var.set("Nothing to save")
            return

        step_number = self._next_step_number()
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"step_{step_number:03d}_{now}.png"
        os.makedirs(SAVE_DIR, exist_ok=True)
        file_path = os.path.join(SAVE_DIR, file_name)

        if preview_window is not None and preview_window.winfo_exists():
            preview_window.destroy()

        self.root.deiconify()
        self.root.lift()

        try:
            image.save(file_path, "PNG")
            self._write_capture_log(step_number, file_name)
            self.status_var.set(f"Saved: {file_path}")
            messagebox.showinfo("Saved", f"Screenshot saved to:\n{file_path}")
            self.preview_state = {}
        except Exception as exc:
            messagebox.showerror("Save Error", f"Failed to save screenshot:\n{exc}")
            self.status_var.set("Save failed")

    def _cancel_preview(self, preview_window):
        if preview_window.winfo_exists():
            preview_window.destroy()
        self.preview_state = {}
        self.root.deiconify()
        self.root.lift()
        self.status_var.set("Save canceled")

    def open_save_folder(self):
        os.makedirs(SAVE_DIR, exist_ok=True)
        try:
            os.startfile(SAVE_DIR)
        except Exception as exc:
            messagebox.showerror("Open Folder Error", f"Could not open folder:\n{exc}")

    def open_progress_frame(self):
        if self.progress_frame_window is not None and self.progress_frame_window.winfo_exists():
            self.progress_frame_window.deiconify()
            self.progress_frame_window.lift()
            self.progress_frame_window.focus_force()
            self.status_var.set("Progress frame ready")
            return

        frame_window = tk.Toplevel(self.root)
        frame_window.title("Progress Camera Frame")
        frame_window.geometry("520x320+120+120")
        frame_window.minsize(240, 160)
        frame_window.attributes("-topmost", True)
        frame_window.configure(bg="#ff00ff")
        self._apply_window_icon(frame_window)

        try:
            frame_window.wm_attributes("-transparentcolor", "#ff00ff")
        except tk.TclError:
            frame_window.attributes("-alpha", 0.35)

        border = tk.Frame(
            frame_window,
            bg="#ff00ff",
            highlightthickness=8,
            highlightbackground="red",
        )
        border.pack(fill="both", expand=True)

        frame_window.protocol("WM_DELETE_WINDOW", self._close_progress_frame)

        for widget in (frame_window, border):
            widget.bind("<ButtonPress-1>", self._progress_frame_on_press)
            widget.bind("<B1-Motion>", self._progress_frame_on_drag)
            widget.bind("<ButtonRelease-1>", self._progress_frame_on_release)
            widget.bind("<Motion>", self._progress_frame_on_motion)
            widget.bind("<Control-MouseWheel>", self._progress_frame_on_ctrl_mousewheel)

        self.progress_frame_window = frame_window
        self.status_var.set("Progress frame created (transparent center)")

    def _close_progress_frame(self):
        self._stop_short_video_capture(open_preview=False)
        if self.progress_frame_window is not None and self.progress_frame_window.winfo_exists():
            self.progress_frame_window.destroy()
        self.progress_frame_window = None
        self.progress_frame_drag_state = {}

    def _progress_frame_get_zone(self, event):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return None

        frame_window = self.progress_frame_window
        width = frame_window.winfo_width()
        height = frame_window.winfo_height()
        if width <= 0 or height <= 0:
            return None

        margin = 36
        x = event.x_root - frame_window.winfo_rootx()
        y = event.y_root - frame_window.winfo_rooty()

        left = x <= margin
        right = x >= (width - margin)
        top = y <= margin
        bottom = y >= (height - margin)

        if top and left:
            return "nw"
        if top and right:
            return "ne"
        if bottom and left:
            return "sw"
        if bottom and right:
            return "se"
        if top:
            return "n"
        if bottom:
            return "s"
        if left:
            return "w"
        if right:
            return "e"
        return "move"

    def _progress_frame_cursor_for_zone(self, zone):
        cursor_map = {
            "nw": "size_nw_se",
            "se": "size_nw_se",
            "ne": "size_ne_sw",
            "sw": "size_ne_sw",
            "n": "size_ns",
            "s": "size_ns",
            "e": "size_we",
            "w": "size_we",
            "move": "fleur",
        }
        return cursor_map.get(zone, "arrow")

    def _progress_frame_on_motion(self, event):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return
        zone = self._progress_frame_get_zone(event)
        self.progress_frame_window.configure(cursor=self._progress_frame_cursor_for_zone(zone))

    def _progress_frame_on_press(self, event):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return

        frame_window = self.progress_frame_window
        frame_window.update_idletasks()

        self.progress_frame_drag_state = {
            "zone": self._progress_frame_get_zone(event) or "move",
            "start_root_x": event.x_root,
            "start_root_y": event.y_root,
            "start_x": frame_window.winfo_x(),
            "start_y": frame_window.winfo_y(),
            "start_w": frame_window.winfo_width(),
            "start_h": frame_window.winfo_height(),
        }

    def _progress_frame_on_drag(self, event):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return

        drag_state = self.progress_frame_drag_state
        if not drag_state:
            return

        zone = drag_state.get("zone")
        dx = event.x_root - drag_state.get("start_root_x", event.x_root)
        dy = event.y_root - drag_state.get("start_root_y", event.y_root)

        min_w = 240
        min_h = 160

        x = drag_state.get("start_x", 0)
        y = drag_state.get("start_y", 0)
        w = drag_state.get("start_w", 240)
        h = drag_state.get("start_h", 160)

        if zone == "move":
            x = drag_state.get("start_x", 0) + dx
            y = drag_state.get("start_y", 0) + dy
        else:
            if "w" in zone:
                x = drag_state.get("start_x", 0) + dx
                w = drag_state.get("start_w", 240) - dx
            if "e" in zone:
                w = drag_state.get("start_w", 240) + dx
            if "n" in zone:
                y = drag_state.get("start_y", 0) + dy
                h = drag_state.get("start_h", 160) - dy
            if "s" in zone:
                h = drag_state.get("start_h", 160) + dy

            if w < min_w:
                if "w" in zone:
                    x -= (min_w - w)
                w = min_w
            if h < min_h:
                if "n" in zone:
                    y -= (min_h - h)
                h = min_h

        self.progress_frame_window.geometry(f"{int(w)}x{int(h)}+{int(x)}+{int(y)}")

    def _progress_frame_on_release(self, _event):
        self.progress_frame_drag_state = {}

    def _progress_frame_on_ctrl_mousewheel(self, event):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return "break"

        frame_window = self.progress_frame_window
        width = frame_window.winfo_width()
        height = frame_window.winfo_height()
        x = frame_window.winfo_x()
        y = frame_window.winfo_y()

        delta = 1 if event.delta > 0 else -1
        step = 24
        new_width = max(240, width + (delta * step))
        new_height = max(160, height + (delta * step))

        dx = (new_width - width) // 2
        dy = (new_height - height) // 2
        frame_window.geometry(f"{new_width}x{new_height}+{x - dx}+{y - dy}")
        return "break"

    def _get_progress_capture_bbox(self):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return None

        frame_window = self.progress_frame_window
        frame_window.update_idletasks()

        x1 = frame_window.winfo_rootx()
        y1 = frame_window.winfo_rooty()
        width = frame_window.winfo_width()
        height = frame_window.winfo_height()

        border_pad = 10
        drag_strip_height = 0
        inner_x1 = x1 + border_pad
        inner_y1 = y1 + border_pad + drag_strip_height
        inner_x2 = x1 + width - border_pad
        inner_y2 = y1 + height - border_pad

        if inner_x2 - inner_x1 < 10 or inner_y2 - inner_y1 < 10:
            return None

        return (inner_x1, inner_y1, inner_x2, inner_y2)

    def capture_progress_frame(self):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            messagebox.showinfo("Progress Frame", "Create the camera frame first.")
            return

        if self.video_capture_running:
            self.status_var.set("Short video capture is running. Click Stop Capture to finish.")
            return

        if self.short_video_mode_armed:
            self._start_short_video_capture()
            return

        bbox = self._get_progress_capture_bbox()
        if bbox is None:
            messagebox.showwarning("Progress Frame", "Frame size is too small to capture.")
            return

        try:
            image = ImageGrab.grab(bbox=bbox, all_screens=True)
            self.progress_frames.append(image.copy())
            self.status_var.set(f"Progress frame captured ({len(self.progress_frames)} total)")
        except Exception as exc:
            messagebox.showerror("Capture Error", f"Could not capture frame:\n{exc}")

    def begin_short_video_mode(self):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            messagebox.showinfo("Short Video", "Create the camera frame first.")
            return

        if self.video_capture_running:
            messagebox.showinfo("Short Video", "Capture is already running. Click Stop Capture first.")
            return

        fps = simpledialog.askinteger(
            "Capture Short Video",
            "Frames per second (FPS):",
            initialvalue=3,
            minvalue=1,
            maxvalue=15,
            parent=self.root,
        )

        if fps is None:
            self.status_var.set("Short video setup canceled")
            return

        self.video_capture_interval_ms = max(int(1000 / fps), 33)
        self.preferred_preview_speed_ms = self.video_capture_interval_ms
        self.short_video_mode_armed = True
        self.status_var.set(f"Short video ready at {fps} FPS. Click Capture Frame to start.")

    def _start_short_video_capture(self):
        if self.video_capture_running:
            return

        bbox = self._get_progress_capture_bbox()
        if bbox is None:
            messagebox.showwarning("Short Video", "Frame size is too small to capture.")
            return

        self.progress_frames = []
        self.video_capture_running = True
        self.short_video_mode_armed = False
        self.status_var.set("Recording short video... Click Stop Capture to finish.")
        self._capture_short_video_tick()

    def _capture_short_video_tick(self):
        if not self.video_capture_running:
            return

        bbox = self._get_progress_capture_bbox()
        if bbox is None:
            self._stop_short_video_capture(open_preview=False)
            messagebox.showwarning("Short Video", "Capture stopped: frame is too small or unavailable.")
            return

        try:
            image = ImageGrab.grab(bbox=bbox, all_screens=True)
            self.progress_frames.append(image.copy())
            self.status_var.set(f"Recording... ({len(self.progress_frames)} frames)")
        except Exception as exc:
            self._stop_short_video_capture(open_preview=False)
            messagebox.showerror("Short Video", f"Capture stopped due to error:\n{exc}")
            return

        self.video_capture_after_id = self.root.after(self.video_capture_interval_ms, self._capture_short_video_tick)

    def _stop_short_video_capture(self, open_preview):
        if self.video_capture_after_id is not None:
            try:
                self.root.after_cancel(self.video_capture_after_id)
            except tk.TclError:
                pass
            self.video_capture_after_id = None

        was_running = self.video_capture_running
        self.video_capture_running = False
        self.short_video_mode_armed = False

        if was_running:
            self.status_var.set(f"Capture stopped ({len(self.progress_frames)} frames)")

        if open_preview and len(self.progress_frames) >= 2:
            self._open_gif_preview_window()

    def stop_short_video_capture(self):
        if not self.video_capture_running:
            if self.short_video_mode_armed:
                self.short_video_mode_armed = False
                self.status_var.set("Short video setup canceled")
            else:
                self.status_var.set("No short video capture running")
            return

        frame_count = len(self.progress_frames)
        self._stop_short_video_capture(open_preview=False)

        if frame_count < 2:
            messagebox.showinfo("Short Video", "Capture stopped. Record at least 2 frames to preview/save GIF.")
            return

        self._open_gif_preview_window()

    def clear_progress_frames(self):
        self._stop_short_video_capture(open_preview=False)
        self.progress_frames = []
        self.status_var.set("Progress frames cleared")

    def export_progress_gif(self):
        if self.video_capture_running:
            messagebox.showinfo("Export GIF", "Stop capture first, then preview/export the GIF.")
            return

        if len(self.progress_frames) < 2:
            messagebox.showinfo("Export GIF", "Capture at least 2 frames to make an animated GIF.")
            return

        self._open_gif_preview_window()

    def _open_gif_preview_window(self):
        if self.gif_preview_state:
            preview_window = self.gif_preview_state.get("window")
            if preview_window is not None and preview_window.winfo_exists():
                preview_window.deiconify()
                preview_window.lift()
                preview_window.focus_force()
                return

        if self.progress_frame_window is not None and self.progress_frame_window.winfo_exists():
            self._close_progress_frame()

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Progress GIF Preview")
        preview_window.transient(self.root)
        preview_window.grab_set()
        self._apply_window_icon(preview_window)

        container = tk.Frame(preview_window, padx=12, pady=12)
        container.pack(fill="both", expand=True)

        tk.Label(
            container,
            text="Preview animation speed before saving your GIF.",
            font=("Segoe UI", 10),
        ).pack(pady=(0, 8))

        preview_label = tk.Label(container)
        preview_label.pack(pady=(0, 8))

        frame_edit_container = tk.Frame(container)
        frame_edit_container.pack(fill="both", pady=(0, 8))

        tk.Label(
            frame_edit_container,
            text="Frames to include:",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

        tk.Label(
            frame_edit_container,
            text="Click thumbnails to keep/remove. Shift+Click selects a range.",
            font=("Segoe UI", 8),
            fg="#444444",
        ).pack(anchor="w", pady=(1, 2))

        frame_scroll_container = tk.Frame(frame_edit_container)
        frame_scroll_container.pack(fill="both", expand=True, pady=(2, 4))

        frame_canvas = tk.Canvas(frame_scroll_container, height=130, highlightthickness=1, highlightbackground="#b0b0b0")
        frame_scrollbar = tk.Scrollbar(frame_scroll_container, orient="horizontal", command=frame_canvas.xview)
        frame_canvas.configure(xscrollcommand=frame_scrollbar.set)

        frame_canvas.pack(fill="both", expand=True)
        frame_scrollbar.pack(fill="x")

        frame_items_container = tk.Frame(frame_canvas)
        frame_canvas.create_window((0, 0), window=frame_items_container, anchor="nw")

        frame_items_container.bind(
            "<Configure>",
            lambda _event: frame_canvas.configure(scrollregion=frame_canvas.bbox("all")),
        )

        selected_count_var = tk.StringVar(value=f"Selected: {len(self.progress_frames)} / {len(self.progress_frames)}")
        frame_select_vars = []
        frame_item_widgets = []
        frame_thumb_refs = []
        preview_photo_refs = []
        fast_resampling = Image.Resampling.BILINEAR if hasattr(Image, "Resampling") else Image.BILINEAR

        for index, frame in enumerate(self.progress_frames):
            row = tk.Frame(frame_items_container, padx=4, pady=2)
            row.pack(side="left", anchor="n")

            thumb_image = frame.copy()
            thumb_image.thumbnail((96, 54), fast_resampling)
            thumb_photo = ImageTk.PhotoImage(thumb_image)
            frame_thumb_refs.append(thumb_photo)

            preview_image = frame.copy()
            preview_image.thumbnail((760, 460), fast_resampling)
            preview_photo = ImageTk.PhotoImage(preview_image)
            preview_photo_refs.append(preview_photo)

            selected_var = tk.BooleanVar(value=True)
            frame_select_vars.append(selected_var)

            frame_btn = tk.Label(
                row,
                text=f"Frame {index + 1}",
                image=thumb_photo,
                compound="top",
                anchor="center",
                padx=6,
                pady=4,
                cursor="hand2",
                bd=2,
                relief="solid",
                highlightthickness=0,
            )
            frame_btn.pack()
            frame_btn.bind(
                "<Button-1>",
                lambda event, frame_index=index: self._on_preview_frame_item_click(frame_index, event),
            )
            frame_item_widgets.append(frame_btn)

        frame_edit_buttons = tk.Frame(frame_edit_container)
        frame_edit_buttons.pack(fill="x")

        select_all_btn = tk.Button(
            frame_edit_buttons,
            text="Select All",
            width=12,
            command=self._select_all_preview_frames,
        )
        select_all_btn.pack(side="left", padx=(0, 6))
        self._add_tooltip(select_all_btn, "Include every captured frame in export.")

        select_none_btn = tk.Button(
            frame_edit_buttons,
            text="Select None",
            width=12,
            command=self._select_none_preview_frames,
        )
        select_none_btn.pack(side="left", padx=(0, 8))
        self._add_tooltip(select_none_btn, "Clear all frame selections.")

        tk.Label(frame_edit_buttons, textvariable=selected_count_var).pack(side="left")

        initial_speed_ms = int(min(max(self.preferred_preview_speed_ms, 80), 2000))
        speed_label_var = tk.StringVar(value=f"Frame delay: {initial_speed_ms} ms")
        speed_label = tk.Label(container, textvariable=speed_label_var)
        speed_label.pack()

        speed_scale = tk.Scale(
            container,
            from_=80,
            to=2000,
            orient="horizontal",
            length=360,
            command=lambda value: self._update_gif_preview_speed(int(float(value))),
        )
        speed_scale.set(initial_speed_ms)
        speed_scale.pack(pady=(2, 8))

        button_row = tk.Frame(container)
        button_row.pack()

        save_btn = tk.Button(button_row, text="Save GIF", width=14, command=self._save_previewed_gif)
        save_btn.pack(side="left", padx=6)
        self._add_tooltip(save_btn, "Save selected frames as an animated GIF (50MB limit).")

        save_mp4_btn = tk.Button(button_row, text="Save MP4", width=14, command=self._save_previewed_mp4)
        save_mp4_btn.pack(side="left", padx=6)
        self._add_tooltip(save_mp4_btn, "Save selected frames as MP4 video.")

        save_webp_btn = tk.Button(button_row, text="Save WebP", width=14, command=self._save_previewed_webp)
        save_webp_btn.pack(side="left", padx=6)
        self._add_tooltip(save_webp_btn, "Save selected frames as animated WebP.")

        cancel_btn = tk.Button(
            button_row,
            text="Cancel",
            width=14,
            command=self._close_gif_preview_window,
        )
        cancel_btn.pack(side="left", padx=6)
        self._add_tooltip(cancel_btn, "Close this preview window.")

        self.gif_preview_state = {
            "window": preview_window,
            "label": preview_label,
            "frames": [],
            "frame_select_vars": frame_select_vars,
            "frame_item_widgets": frame_item_widgets,
            "frame_thumb_refs": frame_thumb_refs,
            "preview_photo_refs": preview_photo_refs,
            "frame_canvas": frame_canvas,
            "selected_count_var": selected_count_var,
            "last_clicked_index": None,
            "index": 0,
            "after_id": None,
            "speed_ms": initial_speed_ms,
            "speed_label_var": speed_label_var,
        }

        frame_canvas.bind("<MouseWheel>", self._on_preview_frame_wheel)
        frame_items_container.bind("<MouseWheel>", self._on_preview_frame_wheel)
        for frame_btn in frame_item_widgets:
            frame_btn.bind("<MouseWheel>", self._on_preview_frame_wheel)

        self._refresh_gif_preview_frames()
        self._animate_gif_preview()
        preview_window.protocol("WM_DELETE_WINDOW", self._close_gif_preview_window)

    def _get_selected_frame_indices(self):
        if not self.gif_preview_state:
            return []
        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        return [index for index, selected_var in enumerate(frame_select_vars) if selected_var.get()]

    def _select_all_preview_frames(self):
        if not self.gif_preview_state:
            return
        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        for selected_var in frame_select_vars:
            selected_var.set(True)
        self.gif_preview_state["last_clicked_index"] = None
        self._refresh_gif_preview_frames()

    def _select_none_preview_frames(self):
        if not self.gif_preview_state:
            return
        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        for selected_var in frame_select_vars:
            selected_var.set(False)
        self.gif_preview_state["last_clicked_index"] = None
        self._refresh_gif_preview_frames()

    def _on_preview_frame_selection(self, _event=None):
        self._refresh_gif_preview_frames()

    def _on_preview_frame_item_click(self, index, event=None):
        if not self.gif_preview_state:
            return

        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        if index < 0 or index >= len(frame_select_vars):
            return

        selected_var = frame_select_vars[index]
        new_state = not selected_var.get()
        shift_pressed = bool(event is not None and (event.state & 0x0001))
        anchor_index = self.gif_preview_state.get("last_clicked_index")

        if shift_pressed and anchor_index is not None and 0 <= anchor_index < len(frame_select_vars):
            start = min(anchor_index, index)
            end = max(anchor_index, index)
            for range_index in range(start, end + 1):
                frame_select_vars[range_index].set(new_state)
        else:
            selected_var.set(new_state)

        self.gif_preview_state["last_clicked_index"] = index
        self._refresh_gif_preview_frames()

    def _refresh_preview_frame_item_styles(self):
        if not self.gif_preview_state:
            return

        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        frame_item_widgets = self.gif_preview_state.get("frame_item_widgets", [])

        for index, widget in enumerate(frame_item_widgets):
            selected = index < len(frame_select_vars) and frame_select_vars[index].get()
            if selected:
                widget.configure(bg="#cfe8ff", relief="solid", bd=2)
            else:
                widget.configure(bg="#f2f2f2", relief="solid", bd=2)

    def _on_preview_frame_wheel(self, event):
        if not self.gif_preview_state:
            return

        frame_canvas = self.gif_preview_state.get("frame_canvas")
        if frame_canvas is None or not frame_canvas.winfo_exists():
            return

        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1

        if delta == 0:
            return

        step = 6 if (event.state & 0x0001) else 3
        frame_canvas.xview_scroll(delta * step, "units")
        return "break"

    def _refresh_gif_preview_frames(self):
        if not self.gif_preview_state:
            return

        selected_indices = self._get_selected_frame_indices()
        selected_count_var = self.gif_preview_state.get("selected_count_var")
        total = len(self.progress_frames)

        if selected_count_var is not None:
            selected_count_var.set(f"Selected: {len(selected_indices)} / {total}")

        self._refresh_preview_frame_item_styles()

        preview_photo_refs = self.gif_preview_state.get("preview_photo_refs", [])
        display_frames = [preview_photo_refs[index] for index in selected_indices if index < len(preview_photo_refs)]

        self.gif_preview_state["frames"] = display_frames
        self.gif_preview_state["index"] = 0

        if not display_frames:
            self.gif_preview_state["label"].configure(image="")

    def _update_gif_preview_speed(self, speed_ms):
        if not self.gif_preview_state:
            return
        self.gif_preview_state["speed_ms"] = speed_ms
        self.preferred_preview_speed_ms = speed_ms
        self.gif_preview_state["speed_label_var"].set(f"Frame delay: {speed_ms} ms")

    def _animate_gif_preview(self):
        if not self.gif_preview_state:
            return

        window = self.gif_preview_state.get("window")
        if window is None or not window.winfo_exists():
            self._close_gif_preview_window()
            return

        frames = self.gif_preview_state["frames"]
        if not frames:
            delay = 200
            after_id = window.after(delay, self._animate_gif_preview)
            self.gif_preview_state["after_id"] = after_id
            return

        label = self.gif_preview_state["label"]
        index = self.gif_preview_state["index"]

        label.configure(image=frames[index])
        next_index = (index + 1) % len(frames)
        self.gif_preview_state["index"] = next_index

        delay = self.gif_preview_state["speed_ms"]
        after_id = window.after(delay, self._animate_gif_preview)
        self.gif_preview_state["after_id"] = after_id

    def _show_save_progress(self, title, maximum, cancel_event=None):
        self._close_save_progress()

        preview_window = None
        if self.gif_preview_state:
            preview_window = self.gif_preview_state.get("window")

        if preview_window is not None and preview_window.winfo_exists():
            parent_window = preview_window
            try:
                preview_window.grab_release()
            except tk.TclError:
                pass
        else:
            parent_window = self.root

        progress_window = tk.Toplevel(self.root)
        progress_window.title(title)
        progress_window.transient(parent_window)
        progress_window.resizable(False, False)
        progress_window.attributes("-topmost", True)
        self._apply_window_icon(progress_window)

        container = tk.Frame(progress_window, padx=14, pady=14)
        container.pack(fill="both", expand=True)

        status_var = tk.StringVar(value="Preparing...")
        status_label = tk.Label(container, textvariable=status_var, justify="left")
        status_label.pack(anchor="w", pady=(0, 8))

        progress = ttk.Progressbar(container, orient="horizontal", mode="determinate", maximum=max(1, maximum), length=340)
        progress.pack(fill="x")

        cancel_button = None
        if cancel_event is not None:
            button_row = tk.Frame(container)
            button_row.pack(fill="x", pady=(10, 0))
            cancel_button = tk.Button(
                button_row,
                text="Cancel",
                width=10,
                command=lambda: self._request_save_cancel(cancel_event),
                font=("Segoe UI", 8),
            )
            cancel_button.pack(side="right")

        self.save_progress_state = {
            "window": progress_window,
            "status_var": status_var,
            "progress": progress,
            "maximum": max(1, maximum),
            "cancel_button": cancel_button,
            "cancel_event": cancel_event,
        }
        self._process_ui_events()

    def _request_save_cancel(self, cancel_event):
        if cancel_event is None or cancel_event.is_set():
            return

        cancel_event.set()
        self._update_save_progress(0, "Cancel requested... stopping export")
        cancel_button = self.save_progress_state.get("cancel_button") if self.save_progress_state else None
        if cancel_button is not None:
            cancel_button.configure(state="disabled")
        self.status_var.set("Canceling export...")

    def _process_ui_events(self):
        try:
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            pass

    def _update_save_progress(self, value, text):
        if not self.save_progress_state:
            return
        status_var = self.save_progress_state.get("status_var")
        progress = self.save_progress_state.get("progress")
        maximum = self.save_progress_state.get("maximum", 1)
        if status_var is not None:
            status_var.set(text)
        if progress is not None:
            progress.configure(value=min(max(value, 0), maximum))
        self._process_ui_events()

    def _close_save_progress(self):
        if not self.save_progress_state:
            return
        window = self.save_progress_state.get("window")
        if window is not None and window.winfo_exists():
            window.destroy()

        if self.gif_preview_state:
            preview_window = self.gif_preview_state.get("window")
            if preview_window is not None and preview_window.winfo_exists():
                try:
                    preview_window.grab_set()
                except tk.TclError:
                    pass
        self.save_progress_state = {}

    def _save_previewed_gif(self):
        if not self.gif_preview_state:
            return

        selected_indices = self._get_selected_frame_indices()
        if len(selected_indices) < 2:
            messagebox.showinfo("Save GIF", "Select at least 2 frames to save a GIF.")
            return

        os.makedirs(SAVE_DIR, exist_ok=True)
        default_name = f"progress_{datetime.now().strftime('%Y%m%d_%H%M%S')}.gif"
        gif_path = filedialog.asksaveasfilename(
            title="Save Progress GIF",
            initialdir=SAVE_DIR,
            initialfile=default_name,
            defaultextension=".gif",
            filetypes=[("GIF", "*.gif")],
        )

        if not gif_path:
            self.status_var.set("GIF export canceled")
            return

        duration_ms = self.gif_preview_state.get("speed_ms", 500)
        selected_frames = [self.progress_frames[index].copy() for index in selected_indices]

        cancel_event = threading.Event()
        self._show_save_progress("Saving GIF", 14, cancel_event=cancel_event)
        self._update_save_progress(1, "Starting GIF export...")
        self.status_var.set("Saving GIF...")

        progress_queue = Queue()
        worker = threading.Thread(
            target=self._run_gif_export_worker,
            args=(selected_frames, duration_ms, gif_path, progress_queue, cancel_event),
            daemon=True,
        )
        worker.start()
        self._poll_gif_export_queue(progress_queue, worker)

    def _run_gif_export_worker(self, selected_frames, duration_ms, gif_path, progress_queue, cancel_event):
        try:
            def progress_callback(attempt, total, text):
                if cancel_event.is_set():
                    raise RuntimeError("CANCELED_BY_USER")
                progress_queue.put(("progress", attempt, text))

            gif_data, used_scale, used_colors, used_frame_count, final_size = self._encode_gif_under_limit(
                selected_frames,
                duration_ms,
                MAX_GIF_BYTES,
                progress_callback=progress_callback,
                cancel_event=cancel_event,
            )

            if cancel_event.is_set():
                progress_queue.put(("canceled", None))
                return

            if gif_data is None:
                progress_queue.put(("over_limit", None))
                return

            progress_queue.put(("progress", 14, "Writing GIF file..."))
            with open(gif_path, "wb") as gif_file:
                gif_file.write(gif_data)

            progress_queue.put(
                (
                    "done",
                    {
                        "gif_path": gif_path,
                        "duration_seconds": (used_frame_count * duration_ms) / 1000.0,
                        "final_size_mb": final_size / (1024 * 1024),
                        "used_scale": used_scale,
                        "used_colors": used_colors,
                        "used_frame_count": used_frame_count,
                    },
                )
            )
        except RuntimeError as exc:
            if str(exc) == "CANCELED_BY_USER":
                progress_queue.put(("canceled", None))
                return
            progress_queue.put(("error", str(exc)))
        except Exception as exc:
            progress_queue.put(("error", str(exc)))

    def _poll_gif_export_queue(self, progress_queue, worker):
        saw_terminal_event = False

        while True:
            try:
                message = progress_queue.get_nowait()
            except Empty:
                break

            event_type = message[0]
            if event_type == "progress":
                _, progress_value, progress_text = message
                self._update_save_progress(progress_value, progress_text)
            elif event_type == "done":
                saw_terminal_event = True
                _, payload = message
                self._close_save_progress()
                gif_path = payload["gif_path"]
                self.status_var.set(f"GIF saved: {gif_path}")
                messagebox.showinfo(
                    "GIF Saved",
                    (
                        f"Progress GIF saved to:\n{gif_path}\n\n"
                        f"Duration: {payload['duration_seconds']:.2f} seconds\n"
                        f"Size: {payload['final_size_mb']:.2f} MB\n"
                        f"Scale: {payload['used_scale']:.2f}x\n"
                        f"Colors: {payload['used_colors']}\n"
                        f"Frames used: {payload['used_frame_count']}"
                    ),
                )
            elif event_type == "over_limit":
                saw_terminal_event = True
                self._close_save_progress()
                messagebox.showerror(
                    "Export Error",
                    "Could not compress GIF under 50MB. Try selecting fewer frames or using a faster frame delay.",
                )
                self.status_var.set("GIF export failed (over 50MB)")
            elif event_type == "canceled":
                saw_terminal_event = True
                self._close_save_progress()
                self.status_var.set("GIF export canceled")
                messagebox.showinfo("Export Canceled", "GIF export was canceled.")
            elif event_type == "error":
                saw_terminal_event = True
                _, error_text = message
                self._close_save_progress()
                messagebox.showerror("Export Error", f"Could not export GIF:\n{error_text}")
                self.status_var.set("GIF export failed")

        if saw_terminal_event:
            return

        if worker.is_alive():
            self.root.after(60, lambda: self._poll_gif_export_queue(progress_queue, worker))
            return

        self._close_save_progress()
        messagebox.showerror(
            "Export Error",
            "GIF export ended unexpectedly. Please try again.",
        )
        self.status_var.set("GIF export failed")

    def _save_previewed_mp4(self):
        if not self.gif_preview_state:
            return

        selected_indices = self._get_selected_frame_indices()
        if len(selected_indices) < 2:
            messagebox.showinfo("Save MP4", "Select at least 2 frames to save an MP4.")
            return

        os.makedirs(SAVE_DIR, exist_ok=True)
        default_name = f"progress_{datetime.now().strftime('%Y%m%d_%H%M%S')}.mp4"
        mp4_path = filedialog.asksaveasfilename(
            title="Save Progress MP4",
            initialdir=SAVE_DIR,
            initialfile=default_name,
            defaultextension=".mp4",
            filetypes=[("MP4", "*.mp4")],
        )

        if not mp4_path:
            self.status_var.set("MP4 export canceled")
            return

        duration_ms = self.gif_preview_state.get("speed_ms", 500)
        fps = max(0.5, 1000.0 / duration_ms)
        selected_frames = [self.progress_frames[index].copy() for index in selected_indices]

        cancel_event = threading.Event()
        self._show_save_progress("Saving MP4", len(selected_frames), cancel_event=cancel_event)
        self._update_save_progress(1, "Starting MP4 export...")
        self.status_var.set("Saving MP4...")

        progress_queue = Queue()
        worker = threading.Thread(
            target=self._run_mp4_export_worker,
            args=(selected_frames, fps, mp4_path, progress_queue, cancel_event),
            daemon=True,
        )
        worker.start()
        self._poll_mp4_export_queue(progress_queue, worker)

    def _run_mp4_export_worker(self, selected_frames, fps, mp4_path, progress_queue, cancel_event):
        try:
            try:
                import imageio.v2 as imageio
                import numpy as np
            except Exception as exc:
                progress_queue.put(("unavailable", str(exc)))
                return

            if cancel_event.is_set():
                progress_queue.put(("canceled", None))
                return

            first_frame = selected_frames[0].convert("RGB")
            base_width, base_height = first_frame.size
            if base_width % 2 != 0:
                base_width -= 1
            if base_height % 2 != 0:
                base_height -= 1
            if base_width <= 0 or base_height <= 0:
                raise ValueError("Frame size is too small for MP4 export")

            resampling = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
            writer = imageio.get_writer(mp4_path, fps=fps, codec="libx264", quality=8)

            try:
                total_frames = len(selected_frames)
                for index, frame in enumerate(selected_frames, start=1):
                    if cancel_event.is_set():
                        raise RuntimeError("CANCELED_BY_USER")
                    rgb_frame = frame.convert("RGB")
                    if rgb_frame.size != (base_width, base_height):
                        rgb_frame = rgb_frame.resize((base_width, base_height), resampling)
                    writer.append_data(np.array(rgb_frame))
                    progress_queue.put(("progress", index, f"Encoding MP4 frame {index}/{total_frames}..."))
            finally:
                writer.close()

            if cancel_event.is_set():
                try:
                    if os.path.exists(mp4_path):
                        os.remove(mp4_path)
                except OSError:
                    pass
                progress_queue.put(("canceled", None))
                return

            progress_queue.put(
                (
                    "done",
                    {
                        "mp4_path": mp4_path,
                        "duration_seconds": len(selected_frames) / fps,
                        "fps": fps,
                        "frames_used": len(selected_frames),
                    },
                )
            )
        except RuntimeError as exc:
            if str(exc) == "CANCELED_BY_USER":
                try:
                    if os.path.exists(mp4_path):
                        os.remove(mp4_path)
                except OSError:
                    pass
                progress_queue.put(("canceled", None))
                return
            progress_queue.put(("error", str(exc)))
        except Exception as exc:
            progress_queue.put(("error", str(exc)))

    def _poll_mp4_export_queue(self, progress_queue, worker):
        saw_terminal_event = False

        while True:
            try:
                message = progress_queue.get_nowait()
            except Empty:
                break

            event_type = message[0]
            if event_type == "progress":
                _, progress_value, progress_text = message
                self._update_save_progress(progress_value, progress_text)
            elif event_type == "done":
                saw_terminal_event = True
                _, payload = message
                self._close_save_progress()
                mp4_path = payload["mp4_path"]
                self.status_var.set(f"MP4 saved: {mp4_path}")
                messagebox.showinfo(
                    "MP4 Saved",
                    (
                        f"Progress MP4 saved to:\n{mp4_path}\n\n"
                        f"Duration: {payload['duration_seconds']:.2f} seconds\n"
                        f"FPS: {payload['fps']:.2f}\n"
                        f"Frames used: {payload['frames_used']}"
                    ),
                )
            elif event_type == "unavailable":
                saw_terminal_event = True
                _, error_text = message
                self._close_save_progress()
                messagebox.showerror(
                    "MP4 Export Unavailable",
                    (
                        "MP4 export dependencies are not available in this build.\n"
                        "Rebuild the app with build_exe.bat and try again.\n\n"
                        f"Details: {error_text}"
                    ),
                )
                self.status_var.set("MP4 export unavailable")
            elif event_type == "canceled":
                saw_terminal_event = True
                self._close_save_progress()
                self.status_var.set("MP4 export canceled")
                messagebox.showinfo("Export Canceled", "MP4 export was canceled.")
            elif event_type == "error":
                saw_terminal_event = True
                _, error_text = message
                self._close_save_progress()
                messagebox.showerror(
                    "Export Error",
                    (
                        f"Could not export MP4:\n{error_text}\n\n"
                        "If this is the first MP4 export, run build_exe.bat again to ensure MP4 dependencies are installed."
                    ),
                )
                self.status_var.set("MP4 export failed")

        if saw_terminal_event:
            return

        if worker.is_alive():
            self.root.after(60, lambda: self._poll_mp4_export_queue(progress_queue, worker))
            return

        self._close_save_progress()
        messagebox.showerror(
            "Export Error",
            "MP4 export ended unexpectedly. Please try again.",
        )
        self.status_var.set("MP4 export failed")

    def _save_previewed_webp(self):
        if not self.gif_preview_state:
            return

        selected_indices = self._get_selected_frame_indices()
        if len(selected_indices) < 2:
            messagebox.showinfo("Save WebP", "Select at least 2 frames to save an animated WebP.")
            return

        os.makedirs(SAVE_DIR, exist_ok=True)
        default_name = f"progress_{datetime.now().strftime('%Y%m%d_%H%M%S')}.webp"
        webp_path = filedialog.asksaveasfilename(
            title="Save Animated WebP",
            initialdir=SAVE_DIR,
            initialfile=default_name,
            defaultextension=".webp",
            filetypes=[("Animated WebP", "*.webp")],
        )

        if not webp_path:
            self.status_var.set("WebP export canceled")
            return

        duration_ms = self.gif_preview_state.get("speed_ms", 500)
        selected_frames = [self.progress_frames[index].copy() for index in selected_indices]

        cancel_event = threading.Event()
        self._show_save_progress("Saving WebP", 1000, cancel_event=cancel_event)
        self._update_save_progress(1, "Starting WebP export...")
        self.status_var.set("Saving WebP...")

        progress_queue = Queue()
        worker = threading.Thread(
            target=self._run_webp_export_worker,
            args=(selected_frames, duration_ms, webp_path, progress_queue, cancel_event),
            daemon=True,
        )
        worker.start()
        self._poll_webp_export_queue(progress_queue, worker)

    def _run_webp_export_worker(self, selected_frames, duration_ms, webp_path, progress_queue, cancel_event):
        try:
            converted_frames = []
            total_frames = len(selected_frames)
            for index, frame in enumerate(selected_frames, start=1):
                if cancel_event.is_set():
                    progress_queue.put(("canceled", None))
                    return
                converted_frames.append(frame.convert("RGB"))
                prep_progress = int((index / total_frames) * 300)
                progress_queue.put(("progress", prep_progress, f"Preparing WebP frame {index}/{total_frames}..."))

            estimated_total_bytes = max(
                800 * 1024,
                int(sum(frame.width * frame.height * 3 for frame in converted_frames) * 0.10),
            )
            last_write_progress = 300

            def on_bytes_written(bytes_written):
                nonlocal estimated_total_bytes, last_write_progress
                if cancel_event.is_set():
                    raise RuntimeError("CANCELED_BY_USER")
                if bytes_written > estimated_total_bytes:
                    estimated_total_bytes = int(bytes_written * 1.25)

                write_ratio = min(bytes_written / max(1, estimated_total_bytes), 0.98)
                write_progress = 300 + int(write_ratio * 650)
                if write_progress > last_write_progress:
                    last_write_progress = write_progress
                    progress_queue.put(
                        (
                            "progress",
                            write_progress,
                            f"Encoding WebP... {bytes_written / (1024 * 1024):.2f} MB written",
                        )
                    )

            progress_queue.put(("progress", 320, "Encoding WebP..."))
            with ProgressFileWriter(
                webp_path,
                progress_callback=on_bytes_written,
                cancel_event=cancel_event,
            ) as monitored_file:
                converted_frames[0].save(
                    monitored_file,
                    format="WEBP",
                    save_all=True,
                    append_images=converted_frames[1:],
                    duration=duration_ms,
                    loop=0,
                    quality=80,
                    method=6,
                )
                on_bytes_written(monitored_file.bytes_written)

            if cancel_event.is_set():
                try:
                    if os.path.exists(webp_path):
                        os.remove(webp_path)
                except OSError:
                    pass
                progress_queue.put(("canceled", None))
                return

            progress_queue.put(("progress", 960, "Finalizing WebP metadata..."))

            with Image.open(webp_path) as check_image:
                webp_frame_count = getattr(check_image, "n_frames", 1)
                is_animated_webp = webp_frame_count > 1

            progress_queue.put(
                (
                    "done",
                    {
                        "webp_path": webp_path,
                        "duration_seconds": (len(converted_frames) * duration_ms) / 1000.0,
                        "frames_used": len(converted_frames),
                        "is_animated_webp": is_animated_webp,
                    },
                )
            )
        except RuntimeError as exc:
            if str(exc) == "CANCELED_BY_USER":
                try:
                    if os.path.exists(webp_path):
                        os.remove(webp_path)
                except OSError:
                    pass
                progress_queue.put(("canceled", None))
                return
            progress_queue.put(("error", str(exc)))
        except Exception as exc:
            progress_queue.put(("error", str(exc)))

    def _poll_webp_export_queue(self, progress_queue, worker):
        saw_terminal_event = False

        while True:
            try:
                message = progress_queue.get_nowait()
            except Empty:
                break

            event_type = message[0]
            if event_type == "progress":
                _, progress_value, progress_text = message
                self._update_save_progress(progress_value, progress_text)
            elif event_type == "done":
                saw_terminal_event = True
                _, payload = message
                self._update_save_progress(1000, "WebP save complete")
                self._close_save_progress()
                webp_path = payload["webp_path"]
                self.status_var.set(f"WebP saved: {webp_path}")
                messagebox.showinfo(
                    "WebP Saved",
                    (
                        f"Animated WebP saved to:\n{webp_path}\n\n"
                        f"Duration: {payload['duration_seconds']:.2f} seconds\n"
                        f"Frames used: {payload['frames_used']}\n"
                        f"Animated file: {'Yes' if payload['is_animated_webp'] else 'No'}\n\n"
                        "Note: Windows Photos/File Explorer often show WebP as static even when animated.\n"
                        "Open the file in Chrome/Edge/Firefox to confirm animation."
                    ),
                )
            elif event_type == "error":
                saw_terminal_event = True
                _, error_text = message
                self._close_save_progress()
                messagebox.showerror(
                    "Export Error",
                    (
                        f"Could not export animated WebP:\n{error_text}\n\n"
                        "If needed, use Save GIF or Save MP4 instead."
                    ),
                )
                self.status_var.set("WebP export failed")
            elif event_type == "canceled":
                saw_terminal_event = True
                self._close_save_progress()
                self.status_var.set("WebP export canceled")
                messagebox.showinfo("Export Canceled", "WebP export was canceled.")

        if saw_terminal_event:
            return

        if worker.is_alive():
            self.root.after(60, lambda: self._poll_webp_export_queue(progress_queue, worker))
            return

        self._close_save_progress()
        messagebox.showerror(
            "Export Error",
            "WebP export ended unexpectedly. Please try again.",
        )
        self.status_var.set("WebP export failed")

    def _encode_gif_under_limit(
        self,
        source_frames,
        duration_ms,
        max_bytes,
        progress_callback=None,
        cancel_event=None,
    ):
        if len(source_frames) < 2:
            return None, None, None, None, None

        resampling = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
        working_frames = list(source_frames)
        scale = 1.0
        colors = 256
        last_size = None
        max_attempts = 14

        for attempt in range(max_attempts):
            if cancel_event is not None and cancel_event.is_set():
                raise RuntimeError("CANCELED_BY_USER")

            if progress_callback is not None:
                progress_callback(
                    attempt + 1,
                    max_attempts,
                    f"Optimizing GIF (attempt {attempt + 1}/{max_attempts})...",
                )

            encoded_frames = []
            for frame_index, frame in enumerate(working_frames, start=1):
                if cancel_event is not None and cancel_event.is_set():
                    raise RuntimeError("CANCELED_BY_USER")

                current = frame.copy()
                if scale < 0.999:
                    resized_width = max(1, int(current.width * scale))
                    resized_height = max(1, int(current.height * scale))
                    current = current.resize((resized_width, resized_height), resampling)
                current = current.convert("P", palette=Image.ADAPTIVE, colors=colors)
                encoded_frames.append(current)
                if frame_index % 3 == 0 and threading.current_thread() is threading.main_thread():
                    self._process_ui_events()

            buffer = io.BytesIO()
            if cancel_event is not None and cancel_event.is_set():
                raise RuntimeError("CANCELED_BY_USER")
            encoded_frames[0].save(
                buffer,
                format="GIF",
                save_all=True,
                append_images=encoded_frames[1:],
                duration=duration_ms,
                loop=0,
                optimize=True,
                disposal=2,
            )

            gif_data = buffer.getvalue()
            gif_size = len(gif_data)
            last_size = gif_size

            if gif_size <= max_bytes:
                return gif_data, scale, colors, len(working_frames), gif_size

            if attempt in {3, 7, 10} and len(working_frames) > 12:
                working_frames = working_frames[::2]
            elif scale > 0.42:
                scale *= 0.86
            elif colors > 64:
                colors = max(64, colors - 32)
            elif len(working_frames) > 2:
                working_frames = working_frames[:-1]
            else:
                break

        return None, scale, colors, len(working_frames), last_size

    def _close_gif_preview_window(self):
        if not self.gif_preview_state:
            return

        window = self.gif_preview_state.get("window")
        after_id = self.gif_preview_state.get("after_id")

        if window is not None and window.winfo_exists() and after_id is not None:
            try:
                window.after_cancel(after_id)
            except tk.TclError:
                pass

        if window is not None and window.winfo_exists():
            window.destroy()

        self.gif_preview_state = {}

    def cancel_snip(self, _event=None):
        if hasattr(self, "overlay") and self.overlay.winfo_exists():
            self.overlay.destroy()
        self.root.deiconify()
        self.root.lift()
        self.status_var.set("Capture canceled")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    SnippingTool().run()
