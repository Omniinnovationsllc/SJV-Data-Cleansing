import tkinter as tk
from tkinter import scrolledtext, messagebox
from threading import Thread
import subprocess
import os
import signal

# Updated list of scripts to run
scripts = [
    "rename.py", "zero.py", "missing.py", "interpol.py", "createfolders.py",
    "timelogrename.py", "tcfilename.py", "validate.py",
    "averaged.py", "original.py",
]

# Globals
current_process = None
mismatch_list = []    # to collect timestamp mismatches across all scripts
error_list = []       # to collect all error lines (tagged "error")

def update_error_counter():
    """Refresh the error counter label."""
    error_label.config(text=f"Errors: {len(error_list)}")

def run_script(script, text_widget):
    """
    Runs a single script, writes output to text_widget,
    highlights errors immediately, and records timestamp mismatches.
    """
    global current_process, mismatch_list, error_list

    current_process = subprocess.Popen(
        ["python", script],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True
    )

    while True:
        line = current_process.stdout.readline()
        if line == "" and current_process.poll() is not None:
            break
        if not line:
            continue

        start_idx = text_widget.index(tk.END)
        text_widget.insert(tk.END, line)
        end_idx = text_widget.index(tk.END)

        lowered = line.lower()

        # 1) Standard error highlight
        if "error" in lowered:
            text_widget.tag_add("error", start_idx, end_idx)
            error_list.append((start_idx, line.strip()))
            update_error_counter()

        # 2) Timestamp‐count mismatch detection
        elif "timestamps" in lowered:
            # Expect format "...: NNNN timestamps"
            parts = line.strip().split()
            try:
                count = int(parts[-2])
            except (ValueError, IndexError):
                count = None

            if count is not None and count != 1440:
                text_widget.tag_add("error", start_idx, end_idx)
                mismatch_list.append(f"{script}: {line.strip()}")
                error_list.append((start_idx, line.strip()))
                update_error_counter()

        text_widget.see(tk.END)

    return current_process.poll()


def run_all_scripts():
    """
    Runs each script in order, tags the "Running" lines, then at the very end
    pops up a single summary of all timestamp mismatches (if any).
    """
    global current_process, mismatch_list, error_list

    mismatch_list.clear()
    error_list.clear()
    update_error_counter()

    for script in scripts:
        # Log start
        start_idx = log_text.index(tk.END)
        log_text.insert(tk.END, f"Running {script}\n")
        end_idx = log_text.index(tk.END)
        log_text.tag_add("running", start_idx, end_idx)
        log_text.see(tk.END)

        # Run and capture exit code
        rc = run_script(script, log_text)

        # Log finish
        log_text.insert(tk.END, f"Finished {script} (exit code {rc})\n")
        log_text.see(tk.END)

    # After all scripts:
    if mismatch_list:
        summary = "\n".join(mismatch_list)
        messagebox.showwarning(
            "Timestamp Mismatch Summary",
            "The following timestamp counts did not equal 1440:\n\n" + summary
        )

    current_process = None


def start_process():
    Thread(target=run_all_scripts, daemon=True).start()


def stop_process():
    """
    Stops the current running process (if any) and logs it.
    """
    global current_process
    if current_process and current_process.poll() is None:
        os.kill(current_process.pid, signal.SIGTERM)
        current_process = None
        log_text.insert(tk.END, "Process stopped by user.\n")
        log_text.see(tk.END)


# --- GUI setup ---
root = tk.Tk()
root.title("Script Runner")

log_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=30)
log_text.pack(padx=10, pady=10)

# Configure tags for coloring
log_text.tag_configure("error", foreground="red")
log_text.tag_configure("running", foreground="#00FF00")  # bright green

button_frame = tk.Frame(root)
button_frame.pack(pady=5)

start_button = tk.Button(button_frame, text="Start", command=start_process)
start_button.pack(side=tk.LEFT, padx=5)

stop_button = tk.Button(button_frame, text="Stop", command=stop_process)
stop_button.pack(side=tk.LEFT, padx=5)

# Error counter label below the buttons
error_label = tk.Label(root, text="Errors: 0", font=("Arial", 12, "bold"))
error_label.pack(pady=(0,10))

root.mainloop()
