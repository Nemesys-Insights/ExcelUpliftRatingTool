

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
from assignment_tool import assign_workbooks  # noqa: F401  <-- EDIT ME




class EvaluatorApp(ttk.Frame):
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

        self.master.title("Evaluator Assignment Generator")
        self.master.minsize(500, 220)

    # ------------------------------------------------------------------
    # UI construction helpers
    # ------------------------------------------------------------------
    def _create_widgets(self):
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.num_evaluators = tk.IntVar(value=3)
        self.eval_per_row = tk.IntVar(value=3)
        self.status_text = tk.StringVar()

        # Input file
        self.lbl_input = ttk.Label(self, text="Input Dataset:")
        self.ent_input = ttk.Entry(self, textvariable=self.input_path)
        self.btn_input = ttk.Button(self, text="Browse…", command=self._browse_input)

        # Output folder
        self.lbl_output = ttk.Label(self, text="Output folder:")
        self.ent_output = ttk.Entry(self, textvariable=self.output_path)
        self.btn_output = ttk.Button(self, text="Browse…", command=self._browse_output)

        # Number of evaluators
        self.lbl_num = ttk.Label(self, text="Number of evaluators (3‒20):")
        self.spn_num = ttk.Spinbox(
            self,
            from_=3,
            to=20,
            textvariable=self.num_evaluators,
            width=5,
        )

         # Number of evaluators per row
        self.lbl_num_eval = ttk.Label(self, text="Evals per response (1-5):")
        self.spn_num_eval = ttk.Spinbox(
            self,
            from_=1,
            to=5,
            textvariable=self.eval_per_row,
            width=5,
        )

        # Status + Run
        self.lbl_status = ttk.Label(self, textvariable=self.status_text, foreground="green")
        self.btn_run = ttk.Button(self, text="Run", command=self._run)

    def _layout_widgets(self):
        self.grid(sticky="nsew")
        pad_y = (0, 25)

        self.lbl_input.grid(row=0, column=0, sticky="w", pady=pad_y)
        self.ent_input.grid(row=0, column=1, sticky="ew", pady=pad_y, padx=(5, 0))
        self.btn_input.grid(row=0, column=2, sticky="e", pady=pad_y, padx=(5, 0))

        self.lbl_output.grid(row=1, column=0, sticky="w", pady=pad_y)
        self.ent_output.grid(row=1, column=1, sticky="ew", pady=pad_y, padx=(5, 0))
        self.btn_output.grid(row=1, column=2, sticky="e", pady=pad_y, padx=(5, 0))

        self.lbl_num.grid(row=2, column=0, sticky="w", pady=pad_y)
        self.spn_num.grid(row=2, column=1, sticky="w", pady=pad_y)

        self.lbl_num_eval.grid(row=3, column=0, sticky ="w", pady=pad_y)
        self.spn_num_eval.grid(row=3, column=1, sticky="w", pady=pad_y)

        self.lbl_status.grid(row=5, column=0, columnspan=3, sticky="w", pady=pad_y)

        self.btn_run.grid(row=6, column=0, columnspan=3, pady=(0, 5))

    # ------------------------------------------------------------------
    # Dialog helpers
    # ------------------------------------------------------------------
    def _browse_input(self):
        filename = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if filename:
            self.input_path.set(filename)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_path.set(folder)

    # ------------------------------------------------------------------
    # Main action
    # ------------------------------------------------------------------
    def _run(self):
        self.status_text.set("Running...")
        self.update()
        input_file = self.input_path.get().strip()
        output_folder = self.output_path.get().strip()

        try:
            num_eval = int(self.num_evaluators.get())
        except (tk.TclError, ValueError):
            num_eval = 0

        try: 
            eval_per_row = int(self.eval_per_row.get())
        except (tk.TclError, ValueError):
            messagebox.showerror("Invalid number of evaluators per row. Defaulting to 3.")
            eval_per_row = 3

        # Basic validation so the user gets instant feedback
        if not input_file:
            messagebox.showerror("Error", "Please choose an input Excel file.")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please choose an output folder.")
            return
        if not (3 <= num_eval <= 20):
            messagebox.showerror("Error", "Number of evaluators must be between 3 and 20.")
            return
        if not (eval_per_row <= num_eval):
            messagebox.showerror("Error", "Number of evaluators must be greater than or equal " \
            "to the desired number of evaluators per row.")
            return 

        # Ensure output directory exists to avoid surprises
        Path(output_folder).mkdir(parents=True, exist_ok=True)

        try:
            assign_workbooks(input_file, output_folder, num_eval)
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
    EvaluatorApp(root)
    sv_ttk.use_dark_theme()
    root.mainloop()
    


if __name__ == "__main__":
    main()
