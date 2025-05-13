import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw
import openpyxl
import os
import datetime

class AnnotationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("H. pylori Annotation Tool")
        self.root.geometry("800x600")
        
        # Apply theme
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure root window
        self.root.configure(bg="#2b2b2b")

        self.frame = ttk.Frame(root, padding="10 5 10 5")
        self.frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar frame
        self.canvas_frame = ttk.Frame(self.frame)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

        self.canvas = tk.Canvas(self.canvas_frame, cursor="cross", bg="white", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")

        self.canvas.configure(xscrollcommand=self.h_scrollbar.set, yscrollcommand=self.v_scrollbar.set)

        self.canvas_frame.rowconfigure(0, weight=1)
        self.canvas_frame.columnconfigure(0, weight=1)

        self.control_frame = ttk.Frame(root, padding="10 5 10 5")
        self.control_frame.pack(fill=tk.X, padx=5, pady=5)

        self.upload_button = ttk.Button(self.control_frame, text="Upload", command=self.upload_images)
        self.upload_button.pack(side=tk.LEFT, padx=(0, 5))

        self.remove_button = ttk.Button(self.control_frame, text="Remove", command=self.remove_last_mark)
        self.remove_button.pack(side=tk.LEFT, padx=(0, 5))

        self.back_button = ttk.Button(self.control_frame, text="Back", command=self.previous_image)
        self.back_button.pack(side=tk.LEFT, padx=(0, 5))

        self.next_button = ttk.Button(self.control_frame, text="Next", command=self.next_image)
        self.next_button.pack(side=tk.LEFT, padx=(0, 5))

        self.save_button = ttk.Button(self.control_frame, text="Save", command=self.save_annotations)
        self.save_button.pack(side=tk.LEFT, padx=(0, 5))

        self.size_label = ttk.Label(self.control_frame, text="Mark Size:", background="#2b2b2b", foreground="white")
        self.size_label.pack(side=tk.LEFT, padx=(0, 5))

        self.size_slider = ttk.Scale(self.control_frame, from_=5, to=100, orient=tk.HORIZONTAL)
        self.size_slider.set(10)
        self.size_slider.pack(side=tk.LEFT, padx=(0, 5))

        self.size_value_label = ttk.Label(self.control_frame, text="10", background="#2b2b2b", foreground="white")
        self.size_value_label.pack(side=tk.LEFT, padx=(0, 5))

        self.size_slider.bind("<Motion>", self.update_size_value)

        self.image_index_label = ttk.Label(self.control_frame, text="", background="#2b2b2b", foreground="white")
        self.image_index_label.pack(side=tk.LEFT, padx=(0, 5))

        self.canvas.bind("<Button-1>", self.mark_position)
        self.images = []
        self.current_image_index = 0
        self.marks = {}
        self.img = None
        self.tk_img = None

        # Add tooltips to buttons
        self.add_tooltip(self.upload_button, "Upload images to annotate")
        self.add_tooltip(self.remove_button, "Remove the last marker")
        self.add_tooltip(self.back_button, "Go to the previous image")
        self.add_tooltip(self.next_button, "Go to the next image")
        self.add_tooltip(self.save_button, "Save annotations to a file")

    def update_size_value(self, event):
        size = self.size_slider.get()
        self.size_value_label.config(text=str(int(size)))

    def add_tooltip(self, widget, text):
        tool_tip = ToolTip(widget, text)

    def upload_images(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg;*.jpeg;*.png")])
        if not files:
            return
        if not self.images:
            self.images = list(files)
        else:
            self.images = self.images[:self.current_image_index + 1] + list(files) + self.images[self.current_image_index + 1:]
        self.load_image()

    def update_image_index_label(self):
        self.image_index_label.config(text=f"Image {self.current_image_index} of {len(self.images) - 1}")

    def load_image(self):
        self.marks[self.current_image_index] = self.marks.get(self.current_image_index, [])
        img_path = self.images[self.current_image_index]
        self.img = Image.open(img_path)
        self.tk_img = ImageTk.PhotoImage(self.img)
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)
        self.update_image_index_label()
        self.redraw_marks()

    def mark_position(self, event):
        x, y = event.x, event.y
        self.marks[self.current_image_index].append((x, y))
        self.draw_mark(x, y)

    def draw_mark(self, x, y):
        size = self.size_slider.get()
        self.canvas.create_rectangle(x - size/2, y - size/2, x + size/2, y + size/2, outline="red")

    def remove_last_mark(self):
        if self.marks[self.current_image_index]:
            self.marks[self.current_image_index].pop()
            self.load_image()

    def redraw_marks(self):
        for mark in self.marks[self.current_image_index]:
            self.draw_mark(mark[0], mark[1])

    def previous_image(self):
        if self.current_image_index > 0:
            self.current_image_index -= 1
            self.load_image()

    def next_image(self):
        if self.current_image_index < len(self.images) - 1:
            self.current_image_index += 1
            self.load_image()

    def save_annotations(self):
        if not self.images:
            messagebox.showwarning("Warning", "No images to save.")
            return

        save_dir = filedialog.askdirectory(title="Select Directory to Save Annotations")
        if not save_dir:
            return

        annotations_file = os.path.join(save_dir, "annotations.xlsx")
        if not os.path.exists(annotations_file):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
        else:
            workbook = openpyxl.load_workbook(annotations_file)
            sheet = workbook.active

        for idx, img_path in enumerate(self.images):
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            img_name = f"image_{idx}_{timestamp}.png"
            annotated_img_path = os.path.join(save_dir, img_name)
            
            # Save the image with the marks
            img = Image.open(img_path)
            draw = ImageDraw.Draw(img)
            size = self.size_slider.get()
            for mark in self.marks.get(idx, []):
                x, y = mark
                draw.rectangle([x - size/2, y - size/2, x + size/2, y + size/2], outline="red")
            img.save(annotated_img_path)

            # Write annotations in the specified format
            sheet.append([f"Image {idx} ({timestamp}):"])
            for mark_idx, mark in enumerate(self.marks.get(idx, [])):
                sheet.append([f"Mark {mark_idx + 1}: {mark}"])

        workbook.save(annotations_file)
        messagebox.showinfo("Info", "Annotations and images saved successfully.")

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event):
        if self.tip_window or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 25
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(tw, text=self.text, background="#ffffe0", relief=tk.SOLID, borderwidth=1)
        label.pack(ipadx=1)

    def hide_tip(self, event):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

if __name__ == "__main__":
    root = tk.Tk()
    app = AnnotationTool(root)
    root.mainloop()
