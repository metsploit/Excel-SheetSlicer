import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading
import os

class SheetSegregator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Sheet Segregator (Fast)")
        self.resizable(False, False)
        self.filename = None
        self.excel = None
        self.df = None

        # --- File picker ---
        file_frame = ttk.Frame(self, padding=10)
        file_frame.grid(row=0, column=0, sticky="ew")
        ttk.Label(file_frame, text="1) Choose workbook:").grid(row=0, column=0, sticky="w")
        self.file_lbl = ttk.Label(file_frame, text="No file selected", width=45)
        self.file_lbl.grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse…", command=self.load_file).grid(row=0, column=2)

        # --- Sheet selector ---
        sheet_frame = ttk.Frame(self, padding=10)
        sheet_frame.grid(row=1, column=0, sticky="ew")
        ttk.Label(sheet_frame, text="2) Select sheet:").grid(row=0, column=0, sticky="w")
        self.sheet_cb = ttk.Combobox(sheet_frame, state="disabled", width=42,
                                     postcommand=self.update_sheets)
        self.sheet_cb.grid(row=0, column=1, padx=5)
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.load_df())

        # --- Column selector ---
        col_frame = ttk.Frame(self, padding=10)
        col_frame.grid(row=2, column=0, sticky="ew")
        ttk.Label(col_frame, text="3) Select column:").grid(row=0, column=0, sticky="w")
        self.col_cb = ttk.Combobox(col_frame, state="disabled", width=42)
        self.col_cb.grid(row=0, column=1, padx=5)

        # --- Progress bar & status ---
        progress_frame = ttk.Frame(self, padding=10)
        progress_frame.grid(row=3, column=0, sticky="ew")
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal",
                                        length=420, mode="determinate")
        self.progress.grid(row=0, column=0, padx=5, pady=(0,4))
        self.status_lbl = ttk.Label(progress_frame, text="Idle")
        self.status_lbl.grid(row=1, column=0, sticky="w")

        # --- Segregate button ---
        btn_frame = ttk.Frame(self, padding=10)
        btn_frame.grid(row=4, column=0, sticky="ew")
        self.go_btn = ttk.Button(btn_frame, text="Segregate (FAST) →", state="disabled",
                                 command=self.start_segregation_thread)
        self.go_btn.grid(row=0, column=1, sticky="e")

        # thread control
        self._seg_thread = None

    def load_file(self):
        fn = filedialog.askopenfilename(
            filetypes=[("Excel files","*.xlsx;*.xlsm;*.xls")])
        if not fn:
            return
        self.filename = fn
        self.file_lbl.config(text=os.path.basename(fn))
        self.sheet_cb.config(state="readonly")
        try:
            self.excel = pd.ExcelFile(self.filename, engine="openpyxl")
            self.sheet_cb['values'] = self.excel.sheet_names
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
            self.sheet_cb.config(state="disabled")

    def update_sheets(self):
        if self.filename:
            try:
                self.excel = pd.ExcelFile(self.filename, engine="openpyxl")
                self.sheet_cb['values'] = self.excel.sheet_names
            except:
                pass

    def load_df(self):
        sheet = self.sheet_cb.get()
        if not sheet:
            return
        try:
            # Read the sheet fully into pandas (fast and enables fast groupby + to_excel)
            self.df = pd.read_excel(self.filename, sheet_name=sheet, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheet:\n{e}")
            return
        cols = list(self.df.columns)
        self.col_cb.config(state="readonly", values=cols)
        self.col_cb.set('')
        self.go_btn.config(state="disabled")
        self.col_cb.bind("<<ComboboxSelected>>", lambda e: self.go_btn.config(state="normal"))

    def start_segregation_thread(self):
        self.go_btn.config(state="disabled")
        self.sheet_cb.config(state="disabled")
        self.col_cb.config(state="disabled")
        self.progress['value'] = 0
        self.status_lbl.config(text="Preparing...")
        self._seg_thread = threading.Thread(target=self.segregate_task, daemon=True)
        self._seg_thread.start()

    def update_progress_ui(self, processed, total, current_sheet_name=None):
        def _update():
            self.progress['maximum'] = max(total, 1)
            self.progress['value'] = processed
            pct = (processed / total * 100) if total > 0 else 0
            status = f"Written {processed}/{total} rows ({pct:.1f}%)"
            if current_sheet_name:
                status += f" — sheet: {current_sheet_name}"
            self.status_lbl.config(text=status)
        self.after(0, _update)

    def finish_ui(self, success=True, message=""):
        def _finish():
            if success:
                self.progress['value'] = self.progress['maximum']
                self.status_lbl.config(text="Done ✅")
                if message:
                    messagebox.showinfo("Done ✅", message)
            else:
                self.status_lbl.config(text="Failed ❌")
                if message:
                    messagebox.showerror("Error", message)
            self.go_btn.config(state="normal")
            self.sheet_cb.config(state="readonly")
            self.col_cb.config(state="readonly")
        self.after(0, _finish)

    @staticmethod
    def _safe_sheet_name(base: str, used: set) -> str:
        """
        Truncate to 31 chars and ensure uniqueness by appending counters if needed.
        """
        base = str(base) if pd.notna(base) else "NaN"
        name = base[:31] if base else "Sheet"
        if not name:
            name = "Sheet"
        original = name
        i = 2
        while name in used:
            suffix = f"_{i}"
            name = (original[:31 - len(suffix)] + suffix) if len(original) + len(suffix) > 31 else original + suffix
            i += 1
        used.add(name)
        return name

    def segregate_task(self):
        col = self.col_cb.get()
        if not col:
            self.finish_ui(success=False, message="Please select a column first.")
            return

        try:
            # Keep only non-null rows for the grouping key
            df_nonnull = self.df[self.df[col].notna()]
            total_rows = len(df_nonnull)
            if total_rows == 0:
                self.finish_ui(success=True, message=f"No values found in '{col}'.")
                return

            # Determine output path: <original>_segregated.xlsx
            base, ext = os.path.splitext(self.filename)
            out_path = f"{base}_segregated.xlsx"

            # Group in order of appearance (sort=False preserves order)
            grouped = df_nonnull.groupby(col, sort=False)

            processed = 0
            self.update_progress_ui(processed, total_rows, None)

            # Ensure xlsxwriter is available
            try:
                import xlsxwriter  # noqa: F401
            except ImportError:
                self.finish_ui(
                    success=False,
                    message="XlsxWriter not installed. Run: pip install XlsxWriter"
                )
                return

            used_names = set()
            # Write all groups to one new workbook with xlsxwriter (fast path)
            # NOTE: removed the 'options' kwarg for compatibility with older pandas
            with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
                # Keep the same column order as original
                cols_order = list(self.df.columns)

                for key, group_df in grouped:
                    # Create a safe/unique sheet name (Excel max 31 chars)
                    sheet_name = self._safe_sheet_name(key, used_names)

                    # Reindex to original column order
                    gd = group_df.reindex(columns=cols_order)

                    # Write the whole group in one shot (header included)
                    gd.to_excel(writer, sheet_name=sheet_name, index=False)

                    # Update progress after each sheet written
                    processed += len(gd)
                    self.update_progress_ui(processed, total_rows, current_sheet_name=sheet_name)

                # Optional: set a basic autofilter on each sheet (cheap)
                for ws_name in used_names:
                    ws = writer.sheets.get(ws_name)
                    if ws is not None and len(cols_order) > 0:
                        ws.autofilter(0, 0, 0, len(cols_order)-1)

            self.finish_ui(
                success=True,
                message=f"Saved segregated workbook:\n{out_path}\n\n"
                        f"Sheets created: {len(used_names)}"
            )

        except Exception as e:
            self.finish_ui(success=False, message=f"Failed:\n{e}")

if __name__ == "__main__":
    app = SheetSegregator()
    app.mainloop()
