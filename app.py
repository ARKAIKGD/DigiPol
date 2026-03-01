import os
import re
import io
import tkinter as tk
from csv import writer
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog, ttk

import imageio.v2 as imageio
import numpy as np
from PIL import Image, ImageDraw, ImageGrab, ImageTk


SAVE_DIR = os.path.join(os.path.expanduser("~"), "Pictures", "StudentSnips")
LOG_PATH = os.path.join(SAVE_DIR, "capture_log.csv")
VERSION_FILE = os.path.join(os.path.dirname(__file__), "version.txt")
MAX_GIF_BYTES = 50 * 1024 * 1024


def get_app_version():
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as version_file:
            value = version_file.read().strip()
            return value if value else "dev"
    except OSError:
        return "dev"


class SnippingTool:
    def __init__(self):
        self.app_version = get_app_version()
        self.root = tk.Tk()
        self.root.title(f"Student Screenshot Tool v{self.app_version}")
        self.root.resizable(False, False)
        self.root.bind_all("<Control-n>", self._on_shortcut_start_snip)
        self.root.bind_all("<Control-o>", self._on_shortcut_open_folder)

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

        self._build_main_ui()
        self._fit_window_to_content()

    def _fit_window_to_content(self):
        self.root.update_idletasks()
        required_width = self.root.winfo_reqwidth()
        required_height = self.root.winfo_reqheight()
        final_width = max(required_width, 400)
        final_height = max(required_height, 260)
        self.root.geometry(f"{final_width}x{final_height}")

    def _build_main_ui(self):
        container = tk.Frame(self.root, padx=14, pady=14)
        container.pack(fill="both", expand=True)

        title = tk.Label(
            container,
            text="Screenshot Snipping Tool",
            font=("Segoe UI", 12, "bold"),
        )
        title.pack(pady=(0, 8))

        description = tk.Label(
            container,
            text="Click Start Snip (Ctrl+N), then drag to capture part of the screen.\nSaves to Pictures\\StudentSnips.",
            wraplength=340,
            justify="center",
        )
        description.pack(pady=(0, 12))

        start_btn = tk.Button(
            container,
            text="Start Snip",
            width=24,
            command=self.start_snip,
            font=("Segoe UI", 10),
        )
        start_btn.pack(pady=(0, 8))

        folder_btn = tk.Button(
            container,
            text="Open Picture Folder (Ctrl+O)",
            width=24,
            command=self.open_save_folder,
            font=("Segoe UI", 10),
        )
        folder_btn.pack(pady=(0, 10))

        progress_label = tk.Label(
            container,
            text="Progress GIF: place frame, capture steps or short video, then export GIF.",
            fg="#333333",
        )
        progress_label.pack(pady=(0, 6))

        progress_row_one = tk.Frame(container)
        progress_row_one.pack(pady=(0, 4))

        frame_btn = tk.Button(
            progress_row_one,
            text="Place Camera Frame",
            width=16,
            command=self.open_progress_frame,
            font=("Segoe UI", 9),
        )
        frame_btn.pack(side="left", padx=4)

        capture_frame_btn = tk.Button(
            progress_row_one,
            text="Capture Frame",
            width=16,
            command=self.capture_progress_frame,
            font=("Segoe UI", 9),
        )
        capture_frame_btn.pack(side="left", padx=4)

        progress_row_two = tk.Frame(container)
        progress_row_two.pack(pady=(0, 8))

        export_gif_btn = tk.Button(
            progress_row_two,
            text="Export Progress GIF",
            width=16,
            command=self.export_progress_gif,
            font=("Segoe UI", 9),
        )
        export_gif_btn.pack(side="left", padx=4)

        clear_frames_btn = tk.Button(
            progress_row_two,
            text="Clear Frames",
            width=16,
            command=self.clear_progress_frames,
            font=("Segoe UI", 9),
        )
        clear_frames_btn.pack(side="left", padx=4)

        progress_row_three = tk.Frame(container)
        progress_row_three.pack(pady=(0, 8))

        short_video_btn = tk.Button(
            progress_row_three,
            text="Capture Short Video",
            width=16,
            command=self.begin_short_video_mode,
            font=("Segoe UI", 9),
        )
        short_video_btn.pack(side="left", padx=4)

        stop_capture_btn = tk.Button(
            progress_row_three,
            text="Stop Capture",
            width=16,
            command=self.stop_short_video_capture,
            font=("Segoe UI", 9),
        )
        stop_capture_btn.pack(side="left", padx=4)

        status = tk.Label(container, textvariable=self.status_var, fg="#333333")
        status.pack()

        version_label = tk.Label(
            container,
            text=f"Version: {self.app_version}",
            fg="#666666",
            font=("Segoe UI", 8),
        )
        version_label.pack(pady=(6, 0))

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

        clear_btn = tk.Button(
            button_row,
            text="Clear",
            width=10,
            command=self._preview_clear,
        )
        clear_btn.pack(side="left", padx=4)

        save_btn = tk.Button(
            button_row,
            text="Save",
            width=14,
            command=lambda: self._save_image(preview_window),
        )
        save_btn.pack(side="left", padx=4)

        cancel_btn = tk.Button(
            button_row,
            text="Cancel",
            width=14,
            command=lambda: self._cancel_preview(preview_window),
        )
        cancel_btn.pack(side="left", padx=4)

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

        try:
            frame_window.wm_attributes("-transparentcolor", "#ff00ff")
        except tk.TclError:
            frame_window.attributes("-alpha", 0.35)

        border = tk.Frame(
            frame_window,
            bg="#ff00ff",
            highlightthickness=3,
            highlightbackground="red",
        )
        border.pack(fill="both", expand=True)

        frame_window.protocol("WM_DELETE_WINDOW", self._close_progress_frame)

        self.progress_frame_window = frame_window
        self.status_var.set("Progress frame created (transparent center)")

    def _close_progress_frame(self):
        self._stop_short_video_capture(open_preview=False)
        if self.progress_frame_window is not None and self.progress_frame_window.winfo_exists():
            self.progress_frame_window.destroy()
        self.progress_frame_window = None

    def _get_progress_capture_bbox(self):
        if self.progress_frame_window is None or not self.progress_frame_window.winfo_exists():
            return None

        frame_window = self.progress_frame_window
        frame_window.update_idletasks()

        x1 = frame_window.winfo_rootx()
        y1 = frame_window.winfo_rooty()
        width = frame_window.winfo_width()
        height = frame_window.winfo_height()

        border_pad = 4
        inner_x1 = x1 + border_pad
        inner_y1 = y1 + border_pad
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

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Progress GIF Preview")
        preview_window.transient(self.root)
        preview_window.grab_set()

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

        frame_scroll_container = tk.Frame(frame_edit_container)
        frame_scroll_container.pack(fill="both", expand=True, pady=(2, 4))

        frame_canvas = tk.Canvas(frame_scroll_container, height=180, highlightthickness=1, highlightbackground="#b0b0b0")
        frame_scrollbar = tk.Scrollbar(frame_scroll_container, orient="vertical", command=frame_canvas.yview)
        frame_canvas.configure(yscrollcommand=frame_scrollbar.set)

        frame_scrollbar.pack(side="right", fill="y")
        frame_canvas.pack(side="left", fill="both", expand=True)

        frame_items_container = tk.Frame(frame_canvas)
        frame_canvas.create_window((0, 0), window=frame_items_container, anchor="nw")

        frame_items_container.bind(
            "<Configure>",
            lambda _event: frame_canvas.configure(scrollregion=frame_canvas.bbox("all")),
        )

        selected_count_var = tk.StringVar(value=f"Selected: {len(self.progress_frames)} / {len(self.progress_frames)}")
        frame_select_vars = []
        frame_thumb_refs = []

        for index, frame in enumerate(self.progress_frames):
            row = tk.Frame(frame_items_container, pady=2)
            row.pack(fill="x", padx=2)

            thumb_image = frame.copy()
            thumb_image.thumbnail((96, 54))
            thumb_photo = ImageTk.PhotoImage(thumb_image)
            frame_thumb_refs.append(thumb_photo)

            selected_var = tk.BooleanVar(value=True)
            frame_select_vars.append(selected_var)

            check_btn = tk.Checkbutton(
                row,
                text=f"Frame {index + 1}",
                variable=selected_var,
                image=thumb_photo,
                compound="left",
                anchor="w",
                command=self._on_preview_frame_selection,
                padx=6,
            )
            check_btn.pack(fill="x")

        frame_edit_buttons = tk.Frame(frame_edit_container)
        frame_edit_buttons.pack(fill="x")

        select_all_btn = tk.Button(
            frame_edit_buttons,
            text="Select All",
            width=12,
            command=self._select_all_preview_frames,
        )
        select_all_btn.pack(side="left", padx=(0, 6))

        select_none_btn = tk.Button(
            frame_edit_buttons,
            text="Select None",
            width=12,
            command=self._select_none_preview_frames,
        )
        select_none_btn.pack(side="left", padx=(0, 8))

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

        save_mp4_btn = tk.Button(button_row, text="Save MP4", width=14, command=self._save_previewed_mp4)
        save_mp4_btn.pack(side="left", padx=6)

        cancel_btn = tk.Button(
            button_row,
            text="Cancel",
            width=14,
            command=self._close_gif_preview_window,
        )
        cancel_btn.pack(side="left", padx=6)

        self.gif_preview_state = {
            "window": preview_window,
            "label": preview_label,
            "frames": [],
            "frame_select_vars": frame_select_vars,
            "frame_thumb_refs": frame_thumb_refs,
            "selected_count_var": selected_count_var,
            "index": 0,
            "after_id": None,
            "speed_ms": initial_speed_ms,
            "speed_label_var": speed_label_var,
        }

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
        self._refresh_gif_preview_frames()

    def _select_none_preview_frames(self):
        if not self.gif_preview_state:
            return
        frame_select_vars = self.gif_preview_state.get("frame_select_vars", [])
        for selected_var in frame_select_vars:
            selected_var.set(False)
        self._refresh_gif_preview_frames()

    def _on_preview_frame_selection(self, _event=None):
        self._refresh_gif_preview_frames()

    def _refresh_gif_preview_frames(self):
        if not self.gif_preview_state:
            return

        selected_indices = self._get_selected_frame_indices()
        selected_count_var = self.gif_preview_state.get("selected_count_var")
        total = len(self.progress_frames)

        if selected_count_var is not None:
            selected_count_var.set(f"Selected: {len(selected_indices)} / {total}")

        display_frames = []
        for index in selected_indices:
            frame = self.progress_frames[index]
            frame_copy = frame.copy()
            frame_copy.thumbnail((760, 460))
            display_frames.append(ImageTk.PhotoImage(frame_copy))

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

    def _show_save_progress(self, title, maximum):
        self._close_save_progress()

        progress_window = tk.Toplevel(self.root)
        progress_window.title(title)
        progress_window.transient(self.root)
        progress_window.grab_set()
        progress_window.resizable(False, False)

        container = tk.Frame(progress_window, padx=14, pady=14)
        container.pack(fill="both", expand=True)

        status_var = tk.StringVar(value="Preparing...")
        status_label = tk.Label(container, textvariable=status_var, justify="left")
        status_label.pack(anchor="w", pady=(0, 8))

        progress = ttk.Progressbar(container, orient="horizontal", mode="determinate", maximum=max(1, maximum), length=340)
        progress.pack(fill="x")

        self.save_progress_state = {
            "window": progress_window,
            "status_var": status_var,
            "progress": progress,
            "maximum": max(1, maximum),
        }
        self.root.update_idletasks()

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
        self.root.update_idletasks()

    def _close_save_progress(self):
        if not self.save_progress_state:
            return
        window = self.save_progress_state.get("window")
        if window is not None and window.winfo_exists():
            window.destroy()
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
        selected_frames = [self.progress_frames[index] for index in selected_indices]

        try:
            self._show_save_progress("Saving GIF", 14)

            def progress_callback(attempt, total, text):
                self._update_save_progress(attempt, text)

            gif_data, used_scale, used_colors, used_frame_count, final_size = self._encode_gif_under_limit(
                selected_frames,
                duration_ms,
                MAX_GIF_BYTES,
                progress_callback=progress_callback,
            )

            if gif_data is None:
                self._close_save_progress()
                messagebox.showerror(
                    "Export Error",
                    "Could not compress GIF under 50MB. Try selecting fewer frames or using a faster frame delay.",
                )
                self.status_var.set("GIF export failed (over 50MB)")
                return

            self._update_save_progress(14, "Writing GIF file...")
            with open(gif_path, "wb") as gif_file:
                gif_file.write(gif_data)

            self._close_save_progress()
            self.status_var.set(f"GIF saved: {gif_path}")
            duration_seconds = (used_frame_count * duration_ms) / 1000.0
            messagebox.showinfo(
                "GIF Saved",
                (
                    f"Progress GIF saved to:\n{gif_path}\n\n"
                    f"Duration: {duration_seconds:.2f} seconds\n"
                    f"Size: {final_size / (1024 * 1024):.2f} MB\n"
                    f"Scale: {used_scale:.2f}x\n"
                    f"Colors: {used_colors}\n"
                    f"Frames used: {used_frame_count}"
                ),
            )
            self._close_gif_preview_window()
        except Exception as exc:
            self._close_save_progress()
            messagebox.showerror("Export Error", f"Could not export GIF:\n{exc}")
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
        selected_frames = [self.progress_frames[index] for index in selected_indices]

        try:
            self._show_save_progress("Saving MP4", len(selected_frames))

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
                for index, frame in enumerate(selected_frames, start=1):
                    rgb_frame = frame.convert("RGB")
                    if rgb_frame.size != (base_width, base_height):
                        rgb_frame = rgb_frame.resize((base_width, base_height), resampling)
                    writer.append_data(np.array(rgb_frame))
                    self._update_save_progress(index, f"Encoding MP4 frame {index}/{len(selected_frames)}...")
            finally:
                writer.close()

            self._close_save_progress()
            duration_seconds = len(selected_frames) / fps
            self.status_var.set(f"MP4 saved: {mp4_path}")
            messagebox.showinfo(
                "MP4 Saved",
                (
                    f"Progress MP4 saved to:\n{mp4_path}\n\n"
                    f"Duration: {duration_seconds:.2f} seconds\n"
                    f"FPS: {fps:.2f}\n"
                    f"Frames used: {len(selected_frames)}"
                ),
            )
        except Exception as exc:
            self._close_save_progress()
            messagebox.showerror(
                "Export Error",
                (
                    f"Could not export MP4:\n{exc}\n\n"
                    "If this is the first MP4 export, run build_exe.bat again to ensure MP4 dependencies are installed."
                ),
            )
            self.status_var.set("MP4 export failed")

    def _encode_gif_under_limit(self, source_frames, duration_ms, max_bytes, progress_callback=None):
        if len(source_frames) < 2:
            return None, None, None, None, None

        resampling = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
        working_frames = list(source_frames)
        scale = 1.0
        colors = 256
        last_size = None
        max_attempts = 14

        for attempt in range(max_attempts):
            if progress_callback is not None:
                progress_callback(
                    attempt + 1,
                    max_attempts,
                    f"Optimizing GIF (attempt {attempt + 1}/{max_attempts})...",
                )

            encoded_frames = []
            for frame in working_frames:
                current = frame.copy()
                if scale < 0.999:
                    resized_width = max(1, int(current.width * scale))
                    resized_height = max(1, int(current.height * scale))
                    current = current.resize((resized_width, resized_height), resampling)
                current = current.convert("P", palette=Image.ADAPTIVE, colors=colors)
                encoded_frames.append(current)

            buffer = io.BytesIO()
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
