import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import subprocess
import threading
import shutil
import logging

# Setup logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')
logger = logging.getLogger(__name__)


# Function to check for updates from central repository
def check_for_updates():
    try:
        subprocess.check_output(['git', 'pull'])  # Adjust command as per your central repository setup
        logger.info("Update check successful.")
    except Exception as e:
        logger.error(f"Error checking for updates: {e}")


class AutomationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Automation Tool")
        self.style = ttk.Style()
        self.setup_styles()
        self.create_widgets()
        self.script_templates = {
            "script1.py": "template1.xlsx",
            "script2.py": "template2.xlsx"
        }
        self.script_requires_excel = {
            "script1.py": False,
            "script2.py": False
        }
        self.uploaded_file_path = None
        self.stop_flag = threading.Event()  # Event to signal stop
        self.process = None  # To store the subprocess object

        # Check for updates when the application starts
        check_for_updates()

    def setup_styles(self):
        self.style.theme_use('clam')
        self.style.configure('TFrame', background='#e0f7fa')
        self.style.configure('TLabel', background='#e0f7fa', font=('Arial', 12))
        self.style.configure('TButton', background='#00796b', foreground='white', font=('Arial', 12, 'bold'))
        self.style.configure('TOptionMenu', background='#00796b', foreground='white', font=('Arial', 12))
        self.style.configure('TProgressbar', background='#00796b')

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="10 10 10 10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        ttk.Label(frame, text="Select Script:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.script_var = tk.StringVar()
        self.script_menu = ttk.OptionMenu(frame, self.script_var, "", *self.get_scripts())
        self.script_menu.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)

        self.download_template_button = ttk.Button(frame, text="Download Template", command=self.download_template)
        self.download_template_button.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

        self.upload_file_button = ttk.Button(frame, text="Upload Excel File", command=self.upload_file)
        self.upload_file_button.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)

        self.run_button = ttk.Button(frame, text="Run", command=self.run_script)
        self.run_button.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)

        self.stop_button = ttk.Button(frame, text="Stop", command=self.stop_script)
        self.stop_button.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)

        self.clear_button = ttk.Button(frame, text="Clear Output", command=self.clear_output)
        self.clear_button.grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)

        self.copy_button = ttk.Button(frame, text="Copy Output", command=self.copy_output)
        self.copy_button.grid(row=0, column=7, padx=5, pady=5, sticky=tk.W)

        self.output_text = ScrolledText(frame, wrap=tk.WORD, width=70, height=20, font=('Arial', 10), bg='#e0f7fa')
        self.output_text.grid(row=1, column=0, columnspan=8, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.progress = ttk.Progressbar(frame, mode='indeterminate')
        self.progress.grid(row=2, column=0, columnspan=8, padx=5, pady=5, sticky=(tk.W, tk.E))

        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.columnconfigure(3, weight=1)
        frame.columnconfigure(4, weight=1)
        frame.columnconfigure(5, weight=1)
        frame.columnconfigure(6, weight=1)
        frame.columnconfigure(7, weight=1)
        frame.rowconfigure(1, weight=1)

    def get_scripts(self):
        script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'automation_scripts')
        if not os.path.exists(script_dir):
            messagebox.showerror("Error", f"Script directory not found: {script_dir}")
            return []
        scripts = [f for f in os.listdir(script_dir) if f.endswith(".py")]
        return scripts

    def download_template(self):
        selected_script = self.script_var.get()
        if not selected_script:
            messagebox.showwarning("Warning", "Please select a script to download its template.")
            return

        template_file = self.script_templates.get(selected_script)
        if not template_file:
            messagebox.showwarning("Warning", "No template available for the selected script.")
            return

        script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'automation_scripts')
        template_path = os.path.join(script_dir, template_file)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if save_path:
            shutil.copyfile(template_path, save_path)
            messagebox.showinfo("Success", f"Template downloaded to {save_path}")

    def upload_file(self):
        self.uploaded_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.uploaded_file_path:
            self.output_text.insert(tk.END, f"Uploaded file: {self.uploaded_file_path}\n")
            self.output_text.see(tk.END)
            self.output_text.update()

    def run_script(self):
        selected_script = self.script_var.get()
        if not selected_script:
            messagebox.showwarning("Warning", "Please select a script to run.")
            return

        if self.script_requires_excel.get(selected_script) and not self.uploaded_file_path:
            messagebox.showwarning("Warning", "Please upload the required Excel file for this script.")
            return

        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'automation_scripts', selected_script)
        self.output_text.insert(tk.END, f"Running {selected_script}...\n")
        self.output_text.see(tk.END)
        self.output_text.update()

        # Start the progress bar animation
        self.progress.start()

        self.stop_flag.clear()  # Clear stop flag before running script

        def execute_script():
            try:
                # Redirect script output to a temporary file
                temp_output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_output.txt')
                with open(temp_output_file, 'w') as f:
                    command = f'python "{script_path}"'
                    if self.uploaded_file_path:
                        command += f' "{self.uploaded_file_path}"'
                    self.process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                                    shell=True)

                    # Read output line by line and log/display it
                    for line in iter(self.process.stdout.readline, b''):
                        line = line.decode("utf-8")
                        logger.info(line.strip())
                        self.root.after(10, self.update_output, line)  # Update GUI

                        # Check if stop flag is set
                        if self.stop_flag.is_set():
                            self.process.terminate()  # Terminate the process if stop flag is set
                            self.output_text.insert(tk.END, "Script stopped by user.\n")
                            self.output_text.see(tk.END)
                            break

                self.process.stdout.close()
                self.process.wait()

            except Exception as e:
                self.output_text.insert(tk.END, f"Error: {e}\n")
            finally:
                # Stop the progress bar animation
                self.progress.stop()
                self.output_text.see(tk.END)

        # Run the script in a separate thread to avoid freezing the GUI
        self.script_thread = threading.Thread(target=execute_script)
        self.script_thread.start()

    def stop_script(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
        self.stop_flag.set()  # Set stop flag to signal the script to stop

    def update_output(self, line):
        """Update the output text widget."""
        self.output_text.insert(tk.END, line + '\n')
        self.output_text.see(tk.END)

    def clear_output(self):
        self.output_text.delete('1.0', tk.END)

    def copy_output(self):
        self.root.clipboard_clear()
        copied_text = self.output_text.get("1.0", tk.END)
        self.root.clipboard_append(copied_text)
        messagebox.showinfo("Copied", "Output copied to clipboard!")


if __name__ == "__main__":
    root = tk.Tk()
    app = AutomationTool(root)
    root.mainloop()
