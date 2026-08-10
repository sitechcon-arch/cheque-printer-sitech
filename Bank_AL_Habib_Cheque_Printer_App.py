import tkinter as tk
from tkinter import ttk, messagebox
import win32print
import win32ui
import win32con

class ChequePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bank AL Habib Cheque Printing Application")
        self.root.geometry("640 x 580")
        self.root.resizable(False, False)

        # Apply Modern Styling
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground="#1F497D")

        # Header Title
        title_label = ttk.Label(root, text="Bank AL Habib - Cheque Writer & Printer", style="Header.TLabel")
        title_label.pack(pady=15)

        # Form Container Frame
        form_frame = ttk.LabelFrame(root, text=" Cheque Information ", padding=15)
        form_frame.pack(fill="x", padx=20, pady=5)

        # 1. Date Field
        ttk.Label(form_frame, text="Date (DDMMYYYY):").grid(row=0, column=0, sticky="w", pady=6)
        self.ent_date = ttk.Entry(form_frame, width=25)
        self.ent_date.insert(0, "10082026")
        self.ent_date.grid(row=0, column=1, sticky="w", pady=6, padx=5)

        # 2. Account Crossing
        ttk.Label(form_frame, text="Account Crossing:").grid(row=1, column=0, sticky="w", pady=6)
        self.ent_crossing = ttk.Entry(form_frame, width=25)
        self.ent_crossing.insert(0, "A/C PAYEE ONLY")
        self.ent_crossing.grid(row=1, column=1, sticky="w", pady=6, padx=5)

        # 3. Payee Name
        ttk.Label(form_frame, text="Payee Name:").grid(row=2, column=0, sticky="w", pady=6)
        self.ent_payee = ttk.Entry(form_frame, width=45)
        self.ent_payee.insert(0, "M/S SITECH CONTRACTING")
        self.ent_payee.grid(row=2, column=1, sticky="w", pady=6, padx=5)

        # 4. Amount in Words Line 1
        ttk.Label(form_frame, text="Amount in Words (Line 1):").grid(row=3, column=0, sticky="w", pady=6)
        self.ent_words1 = ttk.Entry(form_frame, width=45)
        self.ent_words1.insert(0, "One Hundred Thousand Only")
        self.ent_words1.grid(row=3, column=1, sticky="w", pady=6, padx=5)

        # 5. Amount in Words Line 2
        ttk.Label(form_frame, text="Amount in Words (Line 2):").grid(row=4, column=0, sticky="w", pady=6)
        self.ent_words2 = ttk.Entry(form_frame, width=45)
        self.ent_words2.insert(0, "")
        self.ent_words2.grid(row=4, column=1, sticky="w", pady=6, padx=5)

        # 6. Amount in Figures
        ttk.Label(form_frame, text="Amount in Figures (PKR):").grid(row=5, column=0, sticky="w", pady=6)
        self.ent_figures = ttk.Entry(form_frame, width=25)
        self.ent_figures.insert(0, "=100,000/-")
        self.ent_figures.grid(row=5, column=1, sticky="w", pady=6, padx=5)

        # Default Printer Detection
        try:
            default_printer = win32print.GetDefaultPrinter()
        except Exception:
            default_printer = "No default printer detected"

        printer_frame = ttk.Frame(root, padding=10)
        printer_frame.pack(fill="x", padx=20)
        ttk.Label(printer_frame, text=f"🖨️ Selected Printer (Windows Default): {default_printer}", font=("Segoe UI", 9, "italic"), foreground="#444").pack(anchor="w")

        # Action Buttons
        btn_frame = ttk.Frame(root, padding=10)
        btn_frame.pack(fill="x", padx=20, pady=10)

        btn_print = tk.Button(btn_frame, text="🖨️ Send Cheque to Printer", font=("Segoe UI", 11, "bold"), bg="#107C41", fg="white", activebackground="#0B522B", command=self.print_cheque)
        btn_print.pack(fill="x", ipady=8)

    def print_cheque(self):
        try:
            printer_name = win32print.GetDefaultPrinter()
        except Exception as e:
            messagebox.showerror("Printer Error", f"Could not detect standard Windows printer: {e}")
            return

        date_val = self.ent_date.get()
        crossing_val = self.ent_crossing.get()
        payee_val = self.ent_payee.get()
        words1_val = self.ent_words1.get()
        words2_val = self.ent_words2.get()
        figures_val = self.ent_figures.get()

        # Format date for cheque date boxes: e.g. "1 0 0 8 2 0 2 6"
        formatted_date = "  ".join(list(date_val)) if len(date_val) == 8 else date_val

        # Send GDI Print Job directly to Windows Print Spooler
        try:
            hprinter = win32print.OpenPrinter(printer_name)
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            hdc.StartDoc("Bank AL Habib Cheque Print Job")
            hdc.StartPage()

            # Retrieve exact hardware DPI from active printer driver
            dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
            dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)

            # Precise Millimeter to Pixel helper functions
            def mm_to_px_x(mm):
                return int(mm * (dpi_x / 25.4))

            def mm_to_px_y(mm):
                return int(mm * (dpi_y / 25.4))

            # Fonts setup for GDI Device Context
            font_courier = win32ui.CreateFont({
                "name": "Courier New",
                "height": mm_to_px_y(4.5), # Approx 12pt BOLD
                "weight": win32con.FW_BOLD,
            })

            font_arial_small = win32ui.CreateFont({
                "name": "Arial",
                "height": mm_to_px_y(3.2), # Approx 9pt BOLD
                "weight": win32con.FW_BOLD,
            })

            # 1. Account Crossing ("A/C PAYEE ONLY" with parallel lines)
            if crossing_val.strip():
                hdc.SelectObject(font_arial_small)
                pen = win32ui.CreatePen(win32con.PS_SOLID, mm_to_px_y(0.5), 0x000000)
                hdc.SelectObject(pen)
                
                # Parallel top & bottom lines
                hdc.MoveTo(mm_to_px_x(15), mm_to_px_y(5))
                hdc.LineTo(mm_to_px_x(55), mm_to_px_y(5))
                
                hdc.TextOut(mm_to_px_x(18), mm_to_px_y(6), crossing_val)
                
                hdc.MoveTo(mm_to_px_x(15), mm_to_px_y(10))
                hdc.LineTo(mm_to_px_x(55), mm_to_px_y(10))

            # Select Courier font for cheque payload fields
            hdc.SelectObject(font_courier)

            # 2. Date Field (122mm Left, 13.5mm Top)
            hdc.TextOut(mm_to_px_x(122), mm_to_px_y(13.5), formatted_date)

            # 3. Payee Name Field (28mm Left, 25mm Top)
            hdc.TextOut(mm_to_px_x(28), mm_to_px_y(25), payee_val)

            # 4. Amount in Words Line 1 (28mm Left, 34mm Top)
            hdc.TextOut(mm_to_px_x(28), mm_to_px_y(34), words1_val)

            # 5. Amount in Words Line 2 (18mm Left, 42mm Top)
            if words2_val.strip():
                hdc.TextOut(mm_to_px_x(18), mm_to_px_y(42), words2_val)

            # 6. Amount in Figures (122mm Left, 41mm Top)
            hdc.TextOut(mm_to_px_x(122), mm_to_px_y(41), figures_val)

            hdc.EndPage()
            hdc.EndDoc()
            win32print.ClosePrinter(hprinter)

            messagebox.showinfo("Success", f"Cheque sent directly to {printer_name}!")

        except Exception as err:
            messagebox.showerror("Print Job Exception", f"Unable to output print job:
{err}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ChequePrinterApp(root)
    root.mainloop()
