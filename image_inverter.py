"""
Image Inverter - Windows GUI app for image transformations
Supports: color inversion, horizontal flip, vertical flip, rotation
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageOps, ImageTk
import os


class ImageInverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("画像反転ツール")
        self.root.resizable(True, True)

        self.original_image = None
        self.processed_image = None
        self.display_scale = 1.0

        self._build_ui()

    def _build_ui(self):
        # ---- Top toolbar ----
        toolbar = ttk.Frame(self.root, padding=8)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(toolbar, text="画像を開く", command=self.open_image).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="保存", command=self.save_image).pack(side=tk.LEFT, padx=4)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        # Transformation buttons
        ops = [
            ("色反転", self.invert_colors),
            ("左右反転", self.flip_horizontal),
            ("上下反転", self.flip_vertical),
            ("90°回転 →", self.rotate_cw),
            ("90°回転 ←", self.rotate_ccw),
            ("180°回転", self.rotate_180),
        ]
        for label, cmd in ops:
            ttk.Button(toolbar, text=label, command=cmd).pack(side=tk.LEFT, padx=2)

        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)
        ttk.Button(toolbar, text="元に戻す", command=self.reset_image).pack(side=tk.LEFT, padx=4)

        # ---- Status bar ----
        self.status_var = tk.StringVar(value="画像を開いてください")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # ---- Image display area with scrollbars ----
        frame = ttk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(frame, bg="#3c3c3c", cursor="crosshair")
        h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Drag-to-scroll
        self.canvas.bind("<ButtonPress-1>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B1-Motion>", lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))

    # ---- File operations ----

    def open_image(self):
        path = filedialog.askopenfilename(
            title="画像を選択",
            filetypes=[
                ("画像ファイル", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.webp"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path:
            return
        try:
            self.original_image = Image.open(path).convert("RGBA")
            self.processed_image = self.original_image.copy()
            self._show_image()
            name = os.path.basename(path)
            w, h = self.original_image.size
            self.status_var.set(f"{name}  ({w} × {h} px)")
        except Exception as e:
            messagebox.showerror("エラー", f"画像を開けませんでした:\n{e}")

    def save_image(self):
        if self.processed_image is None:
            messagebox.showwarning("警告", "保存する画像がありません")
            return
        path = filedialog.asksaveasfilename(
            title="保存先を選択",
            defaultextension=".png",
            filetypes=[
                ("PNG", "*.png"),
                ("JPEG", "*.jpg"),
                ("BMP", "*.bmp"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path:
            return
        try:
            save_img = self.processed_image
            if path.lower().endswith((".jpg", ".jpeg")):
                save_img = save_img.convert("RGB")
            save_img.save(path)
            self.status_var.set(f"保存しました: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("エラー", f"保存に失敗しました:\n{e}")

    # ---- Transformations ----

    def _require_image(self):
        if self.processed_image is None:
            messagebox.showwarning("警告", "先に画像を開いてください")
            return False
        return True

    def invert_colors(self):
        if not self._require_image():
            return
        r, g, b, a = self.processed_image.split()
        inverted_rgb = ImageOps.invert(Image.merge("RGB", (r, g, b)))
        r2, g2, b2 = inverted_rgb.split()
        self.processed_image = Image.merge("RGBA", (r2, g2, b2, a))
        self._show_image()
        self.status_var.set("色反転 適用済み")

    def flip_horizontal(self):
        if not self._require_image():
            return
        self.processed_image = ImageOps.mirror(self.processed_image)
        self._show_image()
        self.status_var.set("左右反転 適用済み")

    def flip_vertical(self):
        if not self._require_image():
            return
        self.processed_image = ImageOps.flip(self.processed_image)
        self._show_image()
        self.status_var.set("上下反転 適用済み")

    def rotate_cw(self):
        if not self._require_image():
            return
        self.processed_image = self.processed_image.rotate(-90, expand=True)
        self._show_image()
        self.status_var.set("時計回り 90° 回転 適用済み")

    def rotate_ccw(self):
        if not self._require_image():
            return
        self.processed_image = self.processed_image.rotate(90, expand=True)
        self._show_image()
        self.status_var.set("反時計回り 90° 回転 適用済み")

    def rotate_180(self):
        if not self._require_image():
            return
        self.processed_image = self.processed_image.rotate(180, expand=True)
        self._show_image()
        self.status_var.set("180° 回転 適用済み")

    def reset_image(self):
        if self.original_image is None:
            return
        self.processed_image = self.original_image.copy()
        self._show_image()
        self.status_var.set("元の画像に戻しました")

    # ---- Display ----

    def _show_image(self):
        if self.processed_image is None:
            return

        # Fit image to current canvas size (with max scale of 1.0)
        self.canvas.update_idletasks()
        cw = self.canvas.winfo_width() or 800
        ch = self.canvas.winfo_height() or 600
        iw, ih = self.processed_image.size

        scale = min(cw / iw, ch / ih, 1.0)
        display_w = max(1, int(iw * scale))
        display_h = max(1, int(ih * scale))

        display = self.processed_image.resize((display_w, display_h), Image.LANCZOS)
        self._tk_image = ImageTk.PhotoImage(display)

        self.canvas.delete("all")
        self.canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=self._tk_image)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


def main():
    root = tk.Tk()
    root.geometry("900x640")
    app = ImageInverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
