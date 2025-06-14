import tkinter as tk
from tkinter import ttk, messagebox
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import socket

# เปลี่ยน matplotlib เป็นธีมสีขาว
plt.style.use("default")
plt.rcParams['font.family'] = 'Tahoma'
plt.rcParams['axes.facecolor'] = 'white'
plt.rcParams['figure.facecolor'] = 'white'
plt.rcParams['axes.labelcolor'] = 'black'
plt.rcParams['xtick.color'] = 'black'
plt.rcParams['ytick.color'] = 'black'
plt.rcParams['text.color'] = 'black'

class StockSupportCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมวิเคราะห์แนวรับของหุ้น")
        self.root.attributes("-fullscreen", True)
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        self.root.configure(bg="#ffffff")
        self.root.option_add("*Font", ("Tahoma", 13))

        # Apply light theme to ttk
        style = ttk.Style()
        style.theme_use("default")

        # พื้นหลังหลักของแอป
        bg_color = "#f7f9fc"
        fg_color = "#1e1e1e"
        primary_color = "#4a90e2"
        highlight_color = "#e6efff"
        entry_bg = "#ffffff"
        button_bg = "#4a90e2"
        button_fg = "#ffffff"
        button_hover = "#357ab8"

        # พื้นหลังและข้อความปกติ
        style.configure(".", background=bg_color, foreground=fg_color, font=("Tahoma", 12))

        # Label
        style.configure("TLabel", background=bg_color, foreground=fg_color)

        # Entry
        style.configure("TEntry", fieldbackground=entry_bg, foreground=fg_color)

        # Combobox
        style.configure("TCombobox", fieldbackground=entry_bg, foreground=fg_color)

        # Frame
        style.configure("TFrame", background=bg_color)

        # Button
        style.configure("TButton",
            background=button_bg,
            foreground=button_fg,
            padding=6,
            relief="flat"
        )
        style.map("TButton",
            background=[('active', button_hover), ('disabled', '#cccccc')],
            foreground=[('disabled', '#888888')]
        )

        # LabelFrame title
        style.configure("TLabelframe.Label", font=("Tahoma", 12, "bold"), foreground=primary_color)
        style.configure("Small.TLabelframe.Label", font=("Tahoma", 11, "bold"), foreground=primary_color)

        # Labelframe body
        style.configure("TLabelframe", background=bg_color, borderwidth=1, relief="solid")


        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.pack(fill="both", expand=True)

        # --- ส่วนอินพุต ---
        input_frame = ttk.LabelFrame(self.main_frame, text="กรอกข้อมูลหุ้น", padding=15)
        input_frame.pack(fill="x", pady=(0, 15))

        symbol_frame = ttk.Frame(input_frame)
        symbol_frame.pack(fill="x", pady=8)
        ttk.Label(symbol_frame, text="รหัสหุ้น:", width=10).pack(side="left", padx=(0, 10))
        self.symbol_var = tk.StringVar()
        self.symbol_entry = ttk.Entry(symbol_frame, textvariable=self.symbol_var)
        self.symbol_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(symbol_frame, text="(ตัวอย่าง: AAPL, MSFT)", foreground="#555555").pack(side="left", padx=10)

        period_button_frame = ttk.Frame(input_frame)
        period_button_frame.pack(fill="x", pady=8)

        ttk.Label(period_button_frame, text="ช่วงเวลา:", width=10).pack(side="left", padx=(0, 10))
        self.period_var = tk.StringVar(value="3mo")
        period_combo = ttk.Combobox(period_button_frame, textvariable=self.period_var, values=["3mo", "6mo", "1y", "5y"], width=7, state="readonly")
        period_combo.pack(side="left")

        self.calc_button = ttk.Button(period_button_frame, text="คำนวณ", command=self.calculate_support)
        self.calc_button.pack(side="right", padx=5, ipadx=20, ipady=5)

        self.save_pdf_button = ttk.Button(period_button_frame, text="บันทึกเป็น PDF", command=self.save_graph_pdf)
        self.save_pdf_button.pack(side="right", padx=5, ipadx=15, ipady=5)

        # --- ส่วนแสดงผล ---
        # เปลี่ยนข้อความเป็น "ระดับแนวรับ (น้อยไปมาก)"
        self.results_frame = ttk.LabelFrame(self.main_frame, text="ระดับแนวรับ (จากน้อยไปมาก)", padding=15, style="Small.TLabelframe")
        self.results_frame.pack(fill="x", pady=(0, 15))

        support_left_frame = ttk.Frame(self.results_frame)
        support_left_frame.pack(side="left", fill="x", expand=True)

        self.support_labels = []
        # เปลี่ยนข้อความเป็น "แนวรับที่ X:"
        for i in range(3):
            row = ttk.Frame(support_left_frame)
            row.pack(anchor="w", pady=5)
            # แก้ไขข้อความเป็น "แนวรับที่ต่ำสุด", "แนวรับรองลงมา", "แนวรับสูงสุด"
            if i == 0:
                label_text = "แนวรับที่ต่ำสุด:"
            elif i == 1:
                label_text = "แนวรับรองลงมา:"
            else:
                label_text = "แนวรับสูงสุด:"
            ttk.Label(row, text=label_text, width=15, font=("Tahoma", 11)).pack(side="left", padx=5)
            lbl = ttk.Label(row, text="", foreground="#0077cc", font=("Tahoma", 11))
            lbl.pack(side="left")
            self.support_labels.append(lbl)

        self.price_label = ttk.Label(support_left_frame, text="", foreground="#228833", font=("Tahoma", 12, "bold"))
        self.price_label.pack(anchor="w", pady=(10, 0))

        company_frame = ttk.Frame(self.results_frame)
        company_frame.pack(side="right", fill="y")

        self.company_label = ttk.Label(self.results_frame, text="", font=("Tahoma", 14, "bold"), foreground="#0077aa")
        self.company_label.place(relx=0.70, rely=0.4, anchor="center")

        self.status_label = ttk.Label(self.main_frame, text="", foreground="#888888", font=("Tahoma", 10))
        self.status_label.pack(anchor="w", pady=(0, 10))

        # --- กราฟ (ไม่มี Scrollbar แล้ว) ---
        graph_frame = ttk.LabelFrame(self.main_frame, text="กราฟราคาและแนวรับ", padding=10)
        graph_frame.pack(fill="both", expand=True)

        self.figure = plt.Figure(figsize=(8, 4), dpi=100, facecolor="white")
        self.canvas = FigureCanvasTkAgg(self.figure, master=graph_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        self.symbol_entry.bind("<Return>", lambda e: self.calculate_support())

    def has_internet(self):
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=2)
            return True
        except OSError:
            return False

    def validate_symbol(self, symbol):
        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info
            # Check for a common price key to ensure it's a valid, active stock
            return "regularMarketPrice" in info and info["regularMarketPrice"] is not None
        except Exception:
            return False

    def find_support_levels(self, hist):
        close_prices = hist['Close']
        # Find local minima (potential support levels)
        # A point is a local minimum if it's lower than its immediate neighbors
        local_minima = close_prices[(close_prices.shift(1) > close_prices) & (close_prices.shift(-1) > close_prices)]

        # Get the 3 smallest (lowest) support levels and sort them in ascending order
        # This makes the first one the lowest, and the last one the highest of the three.
        support_levels = local_minima.nsmallest(3).sort_values(ascending=True)

        # If we have less than 3 support levels, fill the rest with the historical minimum price
        while len(support_levels) < 3:
            # Append the absolute minimum if we need more levels
            min_price_overall = close_prices.min()
            # Only add if it's not already present or very close to an existing one
            if not any(abs(level - min_price_overall) < 0.01 for level in support_levels):
                support_levels = support_levels.append(pd.Series([min_price_overall]))
            else: # If min_price_overall is already close to an existing support, try a slightly higher value or a different method
                  # For simplicity here, we'll just break or use a placeholder if stuck.
                break # To prevent infinite loop if min is already there
        
        # Ensure we only have 3 levels and they are sorted
        support_levels = support_levels.drop_duplicates().sort_values(ascending=True)[:3]
        
        # If still less than 3, add more placeholders if necessary
        while len(support_levels) < 3:
            # This is a fallback to ensure 3 levels are always shown, even if not true "support"
            # In a real application, you might want more sophisticated logic or fewer displayed levels.
            last_price = close_prices.iloc[-1] if not close_prices.empty else 0
            support_levels = support_levels.append(pd.Series([last_price * (0.95 - (3 - len(support_levels)) * 0.02)])) # Example placeholder
            support_levels = support_levels.sort_values(ascending=True).drop_duplicates()[:3]


        return support_levels.tolist()

    def save_graph_pdf(self):
        from tkinter import filedialog
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if file_path:
            try:
                self.figure.savefig(file_path, format='pdf')
                messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกไฟล์ PDF: {file_path}")
            except Exception as e:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ PDF ได้\n{e}")

    def calculate_support(self):
        symbol = self.symbol_var.get().strip().upper()
        period = self.period_var.get()

        if not symbol:
            messagebox.showwarning("คำเตือน", "กรุณากรอกรหัสหุ้น")
            return

        if not self.has_internet():
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีการเชื่อมต่ออินเทอร์เน็ต")
            return

        # Simple check for valid symbol characters (alphanumeric, no spaces)
        if not symbol.isalnum():
            messagebox.showerror("ข้อผิดพลาด", "รหัสหุ้นไม่ถูกต้อง: ควรเป็นตัวอักษรหรือตัวเลขเท่านั้น")
            return

        self.status_label.config(text="กำลังโหลดข้อมูล...")
        self.root.update()

        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info # Try to get info first to validate
            if not info or "regularMarketPrice" not in info or info["regularMarketPrice"] is None:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่พบข้อมูลหุ้นที่ถูกต้องสำหรับ: {symbol} หรือไม่มีราคาปัจจุบัน")
                self.status_label.config(text="")
                return

            hist = ticker.history(period=period)
            if hist.empty:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่มีข้อมูลประวัติหุ้นสำหรับ {symbol} ในช่วงเวลาที่เลือก")
                self.status_label.config(text="")
                return

            supports = self.find_support_levels(hist)
            # Ensure supports has exactly 3 elements
            while len(supports) < 3:
                # Add a default value if not enough supports are found.
                # This could be a more sophisticated calculation or a fixed low value.
                supports.append(hist['Close'].min() * 0.9) # Example: 90% of the historical min
                supports.sort() # Keep it sorted

            self.company_label.config(text=info.get("longName", symbol)) # Use info directly
            
            # แสดงแนวรับจากน้อยไปมาก
            self.support_labels[0].config(text=f"{supports[0]:.2f} USD") # แนวรับที่ต่ำสุด
            self.support_labels[1].config(text=f"{supports[1]:.2f} USD") # แนวรับรองลงมา
            self.support_labels[2].config(text=f"{supports[2]:.2f} USD") # แนวรับสูงสุด

            self.price_label.config(text=f"ราคาปิดล่าสุด: {hist['Close'][-1]:.2f} USD")
            self.status_label.config(text=f"แสดงข้อมูล {len(hist)} วันย้อนหลัง")

            self.figure.clf()
            ax = self.figure.add_subplot(111)
            ax.plot(hist.index, hist['Close'], label="ราคาปิด", color="#0077cc")
            
            # ลูปเพื่อวาดเส้นแนวรับและเพิ่ม label ที่แตกต่างกัน
            for i, level in enumerate(supports):
                label_text = f"แนวรับที่ต่ำสุด ({level:.2f})" if i == 0 else \
                             f"แนวรับรองลงมา ({level:.2f})" if i == 1 else \
                             f"แนวรับสูงสุด ({level:.2f})"
                ax.axhline(level, color="red", linestyle="--", label=label_text)
            
            ax.set_title(f"กราฟราคาและแนวรับ: {symbol}")
            ax.set_xlabel("วันที่")
            ax.set_ylabel("ราคา (USD)")
            ax.legend()
            self.canvas.draw()

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาดในการดึงข้อมูลหรือคำนวณ\n{e}")
            self.status_label.config(text="")

if __name__ == "__main__":
    root = tk.Tk()
    app = StockSupportCalculator(root)
    root.mainloop()