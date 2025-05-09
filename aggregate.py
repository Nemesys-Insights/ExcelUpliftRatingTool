

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import sys 
import logging

import sv_ttk
# ----------------------------------------------------------------------
# Backend import ‑‑ adjust as needed
# ----------------------------------------------------------------------
# Example: from mypackage.generator import assign_workbooks
from agg_tool import agg_data



class AggregationApp(ttk.Frame):
    """Main application frame."""

    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=20)
        self._build_style()
        self._create_widgets()
        self._layout_widgets()

        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # Style / theme helpers
    # ------------------------------------------------------------------
    def _build_style(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")  # replace with your theme
        except tk.TclError:
            pass

        self.master.title("Evaluation Aggregator")
        self.master.minsize(500, 220)

    # ------------------------------------------------------------------
    # UI construction helpers
    # ------------------------------------------------------------------
    def _create_widgets(self):
        self.input_path = tk.StringVar()
        self.status_text = tk.StringVar()
        
        # Output folder
        self.lbl_input = ttk.Label(self, text="Input folder:")
        self.ent_input = ttk.Entry(self, textvariable=self.input_path)
        self.btn_input = ttk.Button(self, text="Browse…", command=self._browse_input)

        # Status + Run
        self.lbl_status = ttk.Label(self, textvariable=self.status_text, foreground="green")
        self.btn_run = ttk.Button(self, text="Run", command=self._run)

    def _layout_widgets(self):
        self.grid(sticky="nsew")
        pad_y = (0, 25)

        self.lbl_input.grid(row=1, column=0, sticky="w", pady=pad_y)
        self.ent_input.grid(row=1, column=1, sticky="ew", pady=pad_y, padx=(5, 0))
        self.btn_input.grid(row=1, column=2, sticky="e", pady=pad_y, padx=(5, 0))

        self.lbl_status.grid(row=3, column=0, columnspan=3, sticky="w", pady=pad_y)

        self.btn_run.grid(row=4, column=0, columnspan=3, pady=(0, 5))


    def _browse_input(self):
        folder = filedialog.askdirectory(title="Select input folder...")
        if folder:
            self.output_path.set(folder)

    # ------------------------------------------------------------------
    # Main action
    # ------------------------------------------------------------------
    def _run(self):
        self.status_text.set("Running...")
        self.update()
        input_path = self.input_path.get().strip()

      
        # Basic validation so the user gets instant feedback
        if not input_path:
            messagebox.showerror("Error", "Please choose an input Excel file.")
            return
        try:
            result = agg_data(input_path)
            messagebox.showinfo(f"Operation complete. Results saved to {result}")
            self.status_text.set(f"Done.")
            
        except SystemExit as exc:  # validation errors surfaced by backend
            print(logging.error(exc, exc_info=True))
            messagebox.showerror("Validation Error", str(exc))
            self.status_text.set("Failed.")
        except Exception as exc:
            print(logging.error(exc, exc_info=True))
            messagebox.showerror("Processing Error", str(exc))
            self.status_text.set("Failed.")
        else:
            messagebox.showinfo("Success", "Processing completed successfully.")
            self.status_text.set("Done.")


# ----------------------------------------------------------------------
# Script entry point
# ----------------------------------------------------------------------

def main():
    root = tk.Tk()
    AggregationApp(root)
    sv_ttk.use_dark_theme()
    root.mainloop()
    


if __name__ == "__main__":
    main()
