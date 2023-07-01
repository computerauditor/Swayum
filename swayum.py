#© Devyansh Rastogi (@computerauditor)
import pyautogui
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from tkinter import ttk

class MouseCaptureGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Swayum")
        self.root.configure(bg="black")

        self.capture_button = tk.Button(self.root, text="Start Capturing", command=self.start_capturing, fg="yellow",
                                        bg="black")
        self.capture_button.pack()

        self.stop_button = tk.Button(self.root, text="Stop Capturing", command=self.stop_capturing, fg="yellow",
                                     bg="black")
        self.stop_button.pack()
        self.stop_button.config(state=tk.DISABLED)

        self.execute_button = tk.Button(self.root, text="Execute", command=self.execute_positions, fg="yellow",
                                        bg="black")
        self.execute_button.pack()
        self.execute_button.config(state=tk.DISABLED)

        self.save_button = tk.Button(self.root, text="Save Positions", command=self.save_positions, fg="yellow",
                                     bg="black")
        self.save_button.pack()
        self.save_button.config(state=tk.DISABLED)

        self.load_button = tk.Button(self.root, text="Load Positions", command=self.load_positions, fg="yellow",
                                     bg="black")
        self.load_button.pack()
        self.load_button.config(state=tk.NORMAL)

        self.preview_button = tk.Button(self.root, text="Preview Positions", command=self.preview_positions, fg="yellow",
                                        bg="black")
        self.preview_button.pack()
        self.preview_button.config(state=tk.DISABLED)

        self.duration_label = tk.Label(self.root, text="Mouse Speed :", fg="yellow", bg="black")
        self.duration_label.pack()

        self.duration_slider = ttk.Scale(self.root, from_=0, to=10, orient=tk.HORIZONTAL)
        self.duration_slider.pack()

        self.positions = []

        self.root.bind('<Control-i>', self.capture_position)
        self.root.mainloop()

    def start_capturing(self):
        self.capture_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.execute_button.config(state=tk.DISABLED)
        self.save_button.config(state=tk.DISABLED)
        self.load_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.DISABLED)
        self.positions = []
        self.root.bind('<Control-i>', self.capture_position)

    def stop_capturing(self):
        self.capture_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.execute_button.config(state=tk.NORMAL)
        self.save_button.config(state=tk.NORMAL)
        self.load_button.config(state=tk.NORMAL)
        self.preview_button.config(state=tk.NORMAL)
        self.root.unbind('<Control-i>')

    def capture_position(self, event):
        x, y = pyautogui.position()
        self.show_action_dialog(x, y)

    def show_action_dialog(self, x, y):
        action = None

        def select_action_single_click():
            nonlocal action
            action = "Single Click"
            dialog_window.destroy()
            self.capture_action(x, y, action)

        def select_action_right_click():
            nonlocal action
            action = "Right Click"
            dialog_window.destroy()
            self.capture_action(x, y, action)

        def select_action_double_click():
            nonlocal action
            action = "Double Click"
            dialog_window.destroy()
            self.capture_action(x, y, action)

        def select_action_text():
            nonlocal action
            action = "Text"
            dialog_window.destroy()
            self.capture_action(x, y, action)

        dialog_window = tk.Toplevel(self.root)
        dialog_window.title("Select Action")
        dialog_window.configure(bg="black")

        single_click_button = tk.Button(dialog_window, text="Single Click", command=select_action_single_click,
                                        fg="yellow", bg="black")
        single_click_button.pack(pady=10)

        right_click_button = tk.Button(dialog_window, text="Right Click", command=select_action_right_click,
                                       fg="yellow", bg="black")
        right_click_button.pack(pady=10)

        double_click_button = tk.Button(dialog_window, text="Double Click", command=select_action_double_click,
                                        fg="yellow", bg="black")
        double_click_button.pack(pady=10)

        text_button = tk.Button(dialog_window, text="Text", command=select_action_text,
                                fg="yellow", bg="black")
        text_button.pack(pady=10)

        dialog_window.geometry(f"+{x}+{y}")  # Position the dialog at the captured position

    def capture_action(self, x, y, action):
        if action == "Text":
            text = self.enter_text(x, y)
            self.positions.append((x, y, action, text))
            print(f"Position captured: ({x}, {y}), Action: {action}, Text: {text}")
        else:
            self.positions.append((x, y, action))
            print(f"Position captured: ({x}, {y}), Action: {action}")
        self.show_red_dot(x, y)  # Display red dot at the captured position

    def enter_text(self, x, y):
        text = simpledialog.askstring("Enter Text", "Please enter the text:")
        return text

    def execute_positions(self):
        duration = self.duration_slider.get()
        for position in self.positions:
            x, y, action, *extra = position
            if action == "Single Click":
                pyautogui.click(x, y, duration=duration)
            elif action == "Right Click":
                pyautogui.rightClick(x, y, duration=duration)
            elif action == "Double Click":
                pyautogui.doubleClick(x, y, duration=duration)
            elif action == "Text":
                text = extra[0]
                pyautogui.moveTo(x, y, duration=duration)
                pyautogui.typewrite(text, duration=duration)
            print(f"Performed {action} at: ({x}, {y})")

    def show_red_dot(self, x, y):
        dot_size = 10
        dot_color = "red"

        dot_window = tk.Toplevel(self.root)
        dot_window.overrideredirect(True)  # Remove window decorations
        dot_window.attributes("-topmost", True)  # Ensure the dot is on top
        dot_window.attributes("-transparentcolor", dot_color)  # Make the background transparent
        dot_window.configure(bg="black")

        dot_canvas = tk.Canvas(dot_window, width=dot_size, height=dot_size, highlightthickness=0, bg="black")
        dot_canvas.pack()

        dot_canvas.create_oval(0, 0, dot_size, dot_size, fill=dot_color)
        dot_window.geometry(f"+{x - dot_size // 2}+{y - dot_size // 2}")  # Position the dot at the captured position

        # Close the dot window after a delay
        dot_window.after(1000, dot_window.destroy)

    def save_positions(self):
        filename = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if filename:
            with open(filename, "w") as file:
                for position in self.positions:
                    x, y, action, *extra = position
                    if action == "Text":
                        text = extra[0]
                        file.write(f"{x},{y},{action},{text}\n")
                    else:
                        file.write(f"{x},{y},{action}\n")
            print("Positions saved successfully.")

    def load_positions(self):
        filename = tk.filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if filename:
            self.positions = []
            with open(filename, "r") as file:
                for line in file:
                    values = line.strip().split(",")
                    x, y = int(values[0]), int(values[1])
                    action = values[2]
                    if action == "Text":
                        text = values[3]
                        self.positions.append((x, y, action, text))
                    else:
                        self.positions.append((x, y, action))
            print("Positions loaded successfully.")
            self.execute_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
            self.preview_button.config(state=tk.NORMAL)

    def preview_positions(self):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Position Preview")
        preview_window.configure(bg="black")

        canvas_width = self.root.winfo_screenwidth()
        canvas_height = self.root.winfo_screenheight()
        canvas = tk.Canvas(preview_window, width=canvas_width, height=canvas_height, bg="black")
        canvas.pack()

        for i, position in enumerate(self.positions):
            x, y, action, *extra = position
            canvas.create_oval(x - 5, y - 5, x + 5, y + 5, fill="red")
            canvas.create_text(x + 10, y, text=str(i + 1), fill="white", anchor="w")

        preview_window.mainloop()

if __name__ == '__main__':
    gui = MouseCaptureGUI()
