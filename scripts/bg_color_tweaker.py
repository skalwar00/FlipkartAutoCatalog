#!/usr/bin/env python3
"""
bg_color_tweaker.py
-------------------
Tkinter + Pillow GUI app to:
  - Upload multiple images (PNG / JPG / JPEG)
  - Replace near-white / light background pixels with a user-chosen RGB color
    (controlled via real-time sliders)
  - Preview the first uploaded image live
  - Export all processed images to a ZIP with one subfolder per image
"""

import io
import os
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk

# ── constants ─────────────────────────────────────────────────────────────────

CANVAS_W = 600          # preview canvas width  (px)
CANVAS_H = 450          # preview canvas height (px)
BG_TOLERANCE = 40       # how far from white a pixel can be and still be "background"
SLIDER_W = 400          # slider widget width

# ── background detection ──────────────────────────────────────────────────────

def is_background(r, g, b, tolerance=BG_TOLERANCE):
    """
    Return True when the pixel is 'near-white / light background'.
    A pixel qualifies when every channel is within `tolerance` of 255.
    """
    return (255 - r) <= tolerance and (255 - g) <= tolerance and (255 - b) <= tolerance


def replace_background(image: Image.Image, new_r: int, new_g: int, new_b: int,
                        tolerance: int = BG_TOLERANCE) -> Image.Image:
    """
    Return a new PIL image with all background pixels recoloured to (new_r, new_g, new_b).
    Non-background pixels are left unchanged.
    Works on RGBA and RGB source images.
    """
    src = image.convert("RGBA")
    pixels = src.load()
    width, height = src.size

    for y in range(height):
        for x in range(width):
            r, g, b, a = pixels[x, y]
            if a == 0:
                # fully transparent → treat as background
                pixels[x, y] = (new_r, new_g, new_b, 255)
            elif is_background(r, g, b, tolerance):
                pixels[x, y] = (new_r, new_g, new_b, a)

    # return as RGB so JPEG export works too
    return src.convert("RGB")


# ── main application ──────────────────────────────────────────────────────────

class BgColorTweaker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Background Color Tweaker")
        self.resizable(False, False)

        # state
        self._orig_images: list[Image.Image] = []   # original PIL images
        self._image_paths: list[str] = []            # original file paths

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ── top toolbar ───────────────────────────────────────────────────────
        toolbar = tk.Frame(self, bd=1, relief=tk.RIDGE)
        toolbar.pack(fill=tk.X, **pad)

        tk.Button(toolbar, text="Upload Images", width=18,
                  command=self._upload_images).pack(side=tk.LEFT, padx=4, pady=4)

        self._file_label = tk.Label(toolbar, text="No images loaded.",
                                    anchor="w", fg="#555")
        self._file_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)

        tk.Button(toolbar, text="Save All to ZIP", width=18,
                  command=self._save_zip, bg="#2e7d32", fg="white").pack(
                      side=tk.RIGHT, padx=4, pady=4)

        # ── canvas (preview) ──────────────────────────────────────────────────
        canvas_frame = tk.Frame(self, bg="#cccccc", bd=2, relief=tk.SUNKEN)
        canvas_frame.pack(**pad)

        self._canvas = tk.Canvas(canvas_frame, width=CANVAS_W, height=CANVAS_H,
                                 bg="#e0e0e0", cursor="crosshair")
        self._canvas.pack()
        self._canvas_img_ref = None   # keep reference to avoid GC

        # placeholder text
        self._canvas.create_text(CANVAS_W // 2, CANVAS_H // 2,
                                 text="Upload images to preview here",
                                 fill="#999", font=("Helvetica", 14),
                                 tags="placeholder")

        # ── RGB sliders ───────────────────────────────────────────────────────
        slider_frame = tk.LabelFrame(self, text="Background Replacement Color",
                                     padx=10, pady=8)
        slider_frame.pack(fill=tk.X, **pad)

        self._r_var = tk.IntVar(value=255)
        self._g_var = tk.IntVar(value=255)
        self._b_var = tk.IntVar(value=255)

        for label, var, color in (
            ("Red",   self._r_var, "#c0392b"),
            ("Green", self._g_var, "#27ae60"),
            ("Blue",  self._b_var, "#2980b9"),
        ):
            row = tk.Frame(slider_frame)
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=label, width=6, fg=color,
                     font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)

            scale = tk.Scale(row, from_=0, to=255, orient=tk.HORIZONTAL,
                             length=SLIDER_W, variable=var,
                             command=self._on_slider_change,
                             troughcolor=color, highlightthickness=0)
            scale.pack(side=tk.LEFT)

            # live numeric readout
            lbl = tk.Label(row, textvariable=var, width=4,
                           font=("Courier", 10))
            lbl.pack(side=tk.LEFT, padx=4)

        # ── color swatch ──────────────────────────────────────────────────────
        swatch_row = tk.Frame(slider_frame)
        swatch_row.pack(fill=tk.X, pady=(4, 0))
        tk.Label(swatch_row, text="Selected color:").pack(side=tk.LEFT)
        self._swatch = tk.Label(swatch_row, width=8, relief=tk.SOLID, bd=1)
        self._swatch.pack(side=tk.LEFT, padx=6)
        self._update_swatch()

        # ── status bar ────────────────────────────────────────────────────────
        self._status = tk.StringVar(value="Ready.")
        tk.Label(self, textvariable=self._status, anchor="w",
                 fg="#333", font=("Helvetica", 9)).pack(
                     fill=tk.X, side=tk.BOTTOM, padx=10, pady=2)

    # ── event handlers ────────────────────────────────────────────────────────

    def _upload_images(self):
        """Open a file dialog and load the chosen images into memory."""
        paths = filedialog.askopenfilenames(
            title="Select images",
            filetypes=[("Image files", "*.png *.jpg *.jpeg"),
                       ("All files", "*.*")]
        )
        if not paths:
            return

        self._orig_images.clear()
        self._image_paths.clear()

        errors = []
        for p in paths:
            try:
                img = Image.open(p).copy()   # copy so the file handle is released
                self._orig_images.append(img)
                self._image_paths.append(p)
            except Exception as exc:
                errors.append(f"{os.path.basename(p)}: {exc}")

        count = len(self._orig_images)
        self._file_label.config(
            text=f"{count} image(s) loaded." if not errors
            else f"{count} loaded, {len(errors)} failed."
        )
        self._status.set(f"Loaded {count} image(s).")

        if errors:
            messagebox.showwarning("Some images failed to load",
                                   "\n".join(errors))

        if self._orig_images:
            self._update_preview()

    def _on_slider_change(self, _=None):
        """Called whenever any RGB slider moves."""
        self._update_swatch()
        self._update_preview()

    def _update_swatch(self):
        """Redraw the small color swatch to reflect current slider values."""
        hex_color = "#{:02x}{:02x}{:02x}".format(
            self._r_var.get(), self._g_var.get(), self._b_var.get())
        self._swatch.config(bg=hex_color)

    def _update_preview(self):
        """
        Reprocess the *first* loaded image with current slider values and
        display it on the canvas.
        """
        if not self._orig_images:
            return

        r, g, b = self._r_var.get(), self._g_var.get(), self._b_var.get()
        processed = replace_background(self._orig_images[0], r, g, b)

        # fit image inside the canvas while keeping aspect ratio
        processed.thumbnail((CANVAS_W, CANVAS_H), Image.LANCZOS)

        tk_img = ImageTk.PhotoImage(processed)
        self._canvas_img_ref = tk_img          # prevent garbage collection

        self._canvas.delete("all")             # clear placeholder / old image
        # center on canvas
        cx = CANVAS_W // 2
        cy = CANVAS_H // 2
        self._canvas.create_image(cx, cy, anchor=tk.CENTER, image=tk_img)

        self._status.set(
            f"Preview updated — color ({r}, {g}, {b}) — "
            f"{len(self._orig_images)} image(s) queued."
        )

    def _save_zip(self):
        """
        Process every loaded image with the current color settings and pack
        them into a ZIP file chosen by the user.

        ZIP structure:
            <image_name>/
                <image_name>.ext
        """
        if not self._orig_images:
            messagebox.showinfo("No images", "Please upload images first.")
            return

        zip_path = filedialog.asksaveasfilename(
            title="Save ZIP as",
            defaultextension=".zip",
            filetypes=[("ZIP archive", "*.zip")]
        )
        if not zip_path:
            return

        r, g, b = self._r_var.get(), self._g_var.get(), self._b_var.get()
        self._status.set("Processing and saving…")
        self.update_idletasks()

        try:
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for orig_img, orig_path in zip(self._orig_images,
                                               self._image_paths):
                    filename   = os.path.basename(orig_path)
                    name_stem  = os.path.splitext(filename)[0]
                    ext        = os.path.splitext(filename)[1].lower()

                    # determine save format
                    fmt = "PNG" if ext == ".png" else "JPEG"

                    # reprocess with current color
                    processed = replace_background(orig_img, r, g, b)

                    # write to an in-memory buffer — no temp files on disk
                    buf = io.BytesIO()
                    processed.save(buf, format=fmt, quality=95)
                    buf.seek(0)

                    # path inside ZIP:  ImageName/ImageName.ext
                    zip_entry = f"{name_stem}/{filename}"
                    zf.writestr(zip_entry, buf.read())

            self._status.set(
                f"Saved {len(self._orig_images)} image(s) → {os.path.basename(zip_path)}"
            )
            messagebox.showinfo(
                "Export complete",
                f"{len(self._orig_images)} image(s) saved to:\n{zip_path}"
            )

        except Exception as exc:
            self._status.set("Export failed.")
            messagebox.showerror("Export error", str(exc))


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = BgColorTweaker()
    app.mainloop()
