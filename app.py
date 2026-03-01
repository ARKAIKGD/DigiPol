import os
import re
import tkinter as tk
from csv import writer
from datetime import datetime
from tkinter import messagebox

from PIL import ImageDraw, ImageGrab, ImageTk


SAVE_DIR = os.path.join(os.path.expanduser("~"), "Pictures", "StudentSnips")
LOG_PATH = os.path.join(SAVE_DIR, "capture_log.csv")


class SnippingTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Student Screenshot Tool")
        self.root.geometry("360x180")
        self.root.resizable(False, False)
        self.root.bind_all("<Control-n>", self._on_shortcut_start_snip)

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

        self._build_main_ui()

    def _build_main_ui(self):
        container = tk.Frame(self.root, padx=16, pady=16)
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
            wraplength=320,
            justify="center",
        )
        description.pack(pady=(0, 14))

        start_btn = tk.Button(
            container,
            text="Start Snip",
            width=20,
            command=self.start_snip,
            font=("Segoe UI", 10),
        )
        start_btn.pack(pady=(0, 8))

        folder_btn = tk.Button(
            container,
            text="Open Picture Folder",
            width=20,
            command=self.open_save_folder,
            font=("Segoe UI", 10),
        )
        folder_btn.pack(pady=(0, 10))

        status = tk.Label(container, textvariable=self.status_var, fg="#333333")
        status.pack()

    def start_snip(self):
        self.status_var.set("Drag to select an area...")
        self.root.withdraw()
        self.root.after(180, self._open_overlay)

    def _on_shortcut_start_snip(self, _event=None):
        self.start_snip()
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
