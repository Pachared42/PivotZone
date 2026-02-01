import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import socket
from datetime import datetime

plt.style.use("seaborn-v0_8-whitegrid")
plt.rcParams.update({
    "font.family": "Tahoma",
    "axes.facecolor": "#ffffff",
    "figure.facecolor": "#ffffff",
    "axes.edgecolor": "#e5e7eb",
    "axes.labelcolor": "#111827",
    "xtick.color": "#6b7280",
    "ytick.color": "#6b7280",
    "grid.color": "#e5e7eb",
    "grid.alpha": 0.8,
})

class StockSupportCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมวิเคราะห์หุ้น: แนวรับ, แนวต้าน, MA, RSI, Volume")
        
        self.root.attributes("-fullscreen", True)
        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))
        self.root.configure(bg="#ffffff")
        self.root.option_add("*Font", ("Tahoma", 13))

        self._setup_styles() # ตั้งค่า Style สำหรับ ttk widgets
        
        # Initialize instance variables that will hold Tkinter widgets
        # This ensures they exist before _create_main_frame tries to configure them
        self.period_var = tk.StringVar(value="3mo") # ตั้งค่าเริ่มต้นเป็น 3 เดือน
        self.symbol_var = tk.StringVar()
        self.search_history = [] # สำหรับเก็บประวัติการค้นหา

        self.symbol_entry = None # Will be initialized in _create_main_frame
        self.price_label = None
        self.support_labels = []
        self.resistance_labels = []
        self.company_label = None
        self.sector_label = None
        self.industry_label = None
        self.marketcap_label = None
        self.pe_label = None
        self.dividend_label = None
        self.status_label = None
        self.watchlist_listbox = None
        self.calc_button = None
        self.view_graph_button = None
        self.save_pdf_button = None
        self.canvas = None
        self.toolbar = None
        self.figure = None # To hold the matplotlib figure
        self.toolbar_frame = None # To hold the toolbar frame
        self.canvas_widget = None # To hold the canvas widget

        self._setup_main_layout() # ตั้งค่า Layout หลักของหน้าต่าง

        # ตัวแปรสำหรับเก็บข้อมูลที่คำนวณได้ เพื่อใช้ในการสร้างกราฟหรือ PDF
        self.current_hist_data = None
        self.current_support_levels = None
        self.current_resistance_levels = None
        self.current_symbol_info = None

        # จัดการ Watchlist
        self.watchlist_file = "watchlist.txt"
        self.watchlist = self._load_watchlist()
        self._update_watchlist_listbox()

        # จัดการเมื่อผู้ใช้ปิดหน้าต่าง
        root.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _setup_styles(self):
        """
        ตั้งค่า Style สำหรับ ttk widgets เพื่อให้ UI ดูสวยงามและเป็นระเบียบ
        """
        style = ttk.Style()
        style.theme_use("default") # ใช้ Theme เริ่มต้นของระบบ

        # กำหนดสีต่างๆ เพื่อให้ง่ายต่อการแก้ไข
        bg_color = "#f7f9fc" # Background color for most frames
        fg_color = "#1e1e1e" # Foreground color for most text
        primary_color = "#4a90e2" # Primary blue color for accents
        
        entry_bg = "#ffffff" # Background for entry fields
        button_bg = "#4a90e2" # Background for buttons
        button_fg = "#ffffff" # Foreground for button text
        button_hover = "#357ab8" # Button hover effect color
        
        # ตั้งค่า Style ทั่วไป
        style.configure(".", background=bg_color, foreground=fg_color, font=("Tahoma", 12))
        style.configure("TLabel", background=bg_color, foreground=fg_color)
        style.configure("TEntry", fieldbackground=entry_bg, foreground=fg_color)
        style.configure("TCombobox", fieldbackground=entry_bg, foreground=fg_color)
        style.configure("TFrame", background=bg_color)
        
        # Style สำหรับ Button
        style.configure("TButton", background=button_bg, foreground=button_fg, padding=8, relief="flat", font=("Tahoma", 12, "bold"))
        style.map("TButton",
            background=[('active', button_hover), ('disabled', '#cccccc')],
            foreground=[('disabled', '#888888')]
        )
        
        # Style สำหรับ LabelFrame
        style.configure("TLabelframe.Label", font=("Tahoma", 13, "bold"), foreground=primary_color)
        style.configure("Small.TLabelframe.Label", font=("Tahoma", 12, "bold"), foreground=primary_color)
        style.configure("TLabelframe", background=bg_color, borderwidth=1, relief="solid")

    def _setup_main_layout(self):
        """
        ตั้งค่า Layout หลักของหน้าต่าง และสร้าง Frames สำหรับแต่ละหน้าจอ
        """
        # กำหนดให้ root grid row/column 0 ขยายเต็ม
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Main container frame: ใช้สำหรับสลับหน้าจอหลักกับหน้าจอกราฟ
        self.container_frame = ttk.Frame(self.root)
        self.container_frame.grid(row=0, column=0, sticky="nsew")
        self.container_frame.grid_rowconfigure(0, weight=1)
        self.container_frame.grid_columnconfigure(0, weight=1)

        self.frames = {} # Dictionary สำหรับเก็บ frame ของแต่ละหน้า

        self._create_main_frame() # สร้างหน้าจอหลัก
        self._create_graph_frame() # สร้างหน้าจอกราฟ

        self.show_frame("main") # แสดงหน้าจอหลักเป็นอันดับแรก

        # ผูกปุ่ม Enter กับฟังก์ชันคำนวณในช่องกรอกหุ้น
        self.symbol_entry.bind("<Return>", lambda e: self.calculate_support())
    
    def _create_main_frame(self):
        """
        สร้างและจัดวางองค์ประกอบ UI สำหรับหน้าจอหลัก
        """
        main_frame = ttk.Frame(self.container_frame, padding=25)
        self.frames["main"] = main_frame
        
        # กำหนด weight ให้ rows และ columns ที่ต้องการให้ขยายตัว
        main_frame.grid_rowconfigure(1, weight=1) # Row 1 สำหรับ results_frame และ watchlist_frame
        main_frame.grid_columnconfigure(0, weight=3) # Column 0 สำหรับ input_frame, results_frame, status_label
        main_frame.grid_columnconfigure(1, weight=1) # Column 1 สำหรับ watchlist_frame

        # --- กรอบสำหรับ Input ข้อมูลหุ้น ---
        input_frame = ttk.LabelFrame(main_frame, text="กรอกข้อมูลหุ้น", padding=20)
        input_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))

        symbol_row_frame = ttk.Frame(input_frame)
        symbol_row_frame.pack(fill="x", pady=8)
        ttk.Label(symbol_row_frame, text="รหัสหุ้น:", width=10).pack(side="left", padx=(0, 10))
        self.symbol_entry = ttk.Combobox(symbol_row_frame, textvariable=self.symbol_var, values=self.search_history, font=("Tahoma", 13))
        self.symbol_entry.pack(side="left", fill="x", expand=True)
        ttk.Label(symbol_row_frame, text="(ตัวอย่าง: AAPL, MSFT)", foreground="#555555").pack(side="left", padx=10)

        period_button_row_frame = ttk.Frame(input_frame)
        period_button_row_frame.pack(fill="x", pady=8)

        ttk.Label(period_button_row_frame, text="ช่วงเวลา:", width=10).pack(side="left", padx=(0, 10))
        self.period_combo = ttk.Combobox(period_button_row_frame, textvariable=self.period_var, 
                                          values=["3mo", "6mo", "1y", "5y", "10y", "max"], 
                                          width=7, state="readonly", font=("Tahoma", 13))
        self.period_combo.pack(side="left", padx=(0, 20))

        self.save_pdf_button = ttk.Button(period_button_row_frame, text="บันทึกกราฟเป็น PDF", command=self.save_graph_pdf)
        self.save_pdf_button.pack(side="right", padx=5)
        self.save_pdf_button.config(state="disabled")

        self.view_graph_button = ttk.Button(period_button_row_frame, text="ดูกราฟ", command=self.show_graph_frame)
        self.view_graph_button.pack(side="right", padx=5)
        self.view_graph_button.config(state="disabled")

        self.calc_button = ttk.Button(period_button_row_frame, text="คำนวณ", command=self.calculate_support)
        self.calc_button.pack(side="right", padx=5)

        # --- กรอบสำหรับแสดงผลข้อมูลสำคัญของหุ้น ---
        self.results_frame = ttk.LabelFrame(main_frame, text="ข้อมูลสำคัญและระดับราคา", padding=20, style="Small.TLabelframe")
        self.results_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 15), pady=(0, 10))
        self.results_frame.grid_columnconfigure(0, weight=1) # ให้คอลัมน์ด้านซ้ายขยาย (price_info_frame)
        self.results_frame.grid_columnconfigure(1, weight=1) # ให้คอลัมน์ด้านขวาขยาย (company_info_frame)
        self.results_frame.grid_rowconfigure(0, weight=1) # ให้แถวภายในขยายตาม

        # เฟรมสำหรับแสดงราคาปัจจุบัน แนวรับ แนวต้าน
        price_info_frame = ttk.Frame(self.results_frame)
        price_info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.price_label = ttk.Label(price_info_frame, text="", foreground="#228833", font=("Tahoma", 16, "bold"))
        self.price_label.pack(anchor="w", pady=(5, 15))

        ttk.Label(price_info_frame, text="ระดับแนวรับ:", font=("Tahoma", 12, "bold"), foreground="#4a90e2").pack(anchor="w", pady=(5, 5))
        self.support_labels = []
        support_texts = ["ต่ำสุด:", "รองลงมา:", "สูงสุด:"]
        for i in range(3):
            row = ttk.Frame(price_info_frame)
            row.pack(anchor="w", pady=2)
            ttk.Label(row, text=f"แนวรับ {support_texts[i]}", width=12, font=("Tahoma", 12)).pack(side="left", padx=5)
            lbl = ttk.Label(row, text="", foreground="#0077cc", font=("Tahoma", 12, "bold"))
            lbl.pack(side="left")
            self.support_labels.append(lbl)

        ttk.Label(price_info_frame, text="ระดับแนวต้าน:", font=("Tahoma", 12, "bold"), foreground="#4a90e2").pack(anchor="w", pady=(15, 5))
        self.resistance_labels = []
        resistance_texts = ["ต่ำสุด:", "รองลงมา:", "สูงสุด:"]
        for i in range(3):
            row = ttk.Frame(price_info_frame)
            row.pack(anchor="w", pady=2)
            ttk.Label(row, text=f"แนวต้าน {resistance_texts[i]}", width=12, font=("Tahoma", 12)).pack(side="left", padx=5)
            lbl = ttk.Label(row, text="", foreground="#cc0077", font=("Tahoma", 12, "bold"))
            lbl.pack(side="left")
            self.resistance_labels.append(lbl)
            
        # เฟรมสำหรับแสดงข้อมูลบริษัท
        company_info_frame = ttk.Frame(self.results_frame)
        company_info_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.company_label = ttk.Label(company_info_frame, text="", font=("Tahoma", 18, "bold"), foreground="#4a90e2")
        self.company_label.pack(anchor="w", pady=(0, 15))
        
        # Corrected usage: Assign the created label directly to the instance variable
        self.sector_label = self._add_info_label(company_info_frame, "ภาคส่วน:")
        self.industry_label = self._add_info_label(company_info_frame, "อุตสาหกรรม:")
        self.marketcap_label = self._add_info_label(company_info_frame, "มูลค่าตลาด:")
        self.pe_label = self._add_info_label(company_info_frame, "P/E Ratio (TTM):")
        self.dividend_label = self._add_info_label(company_info_frame, "อัตราเงินปันผล (%):")

        # Label แสดงสถานะ
        self.status_label = ttk.Label(main_frame, text="", foreground="#888888", font=("Tahoma", 11))
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0), padx=(0, 10))

        # --- กรอบสำหรับ Watchlist ---
        watchlist_frame = ttk.LabelFrame(main_frame, text="รายการหุ้นโปรด", padding=15, style="Small.TLabelframe")
        watchlist_frame.grid(row=1, column=1, sticky="nsew", padx=(15, 0), pady=(0, 10))

        self.watchlist_listbox = tk.Listbox(watchlist_frame, height=10, font=("Tahoma", 12), selectmode=tk.SINGLE, 
                                             bg="#ffffff", fg="#1e1e1e", selectbackground="#4a90e2", selectforeground="#ffffff")
        self.watchlist_listbox.pack(fill="both", expand=True, pady=(5, 0))
        
        watchlist_buttons_frame = ttk.Frame(watchlist_frame)
        watchlist_buttons_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(watchlist_buttons_frame, text="เพิ่ม", command=self._add_to_watchlist).pack(side="left", expand=True, padx=2)
        ttk.Button(watchlist_buttons_frame, text="ลบ", command=self._remove_from_watchlist).pack(side="right", expand=True, padx=2)
        
        self.watchlist_listbox.bind("<<ListboxSelect>>", self._on_watchlist_select)

    def _add_info_label(self, parent_frame, label_text):
        """
        สร้าง Label คู่ (ข้อความหัวข้อ + ข้อความข้อมูล) และคืนค่า Label ที่แสดงข้อมูล
        """
        frame = ttk.Frame(parent_frame)
        frame.pack(anchor="w", pady=2)
        ttk.Label(frame, text=label_text, font=("Tahoma", 12)).pack(side="left")
        var_label = ttk.Label(frame, text="", foreground="#1e1e1e", font=("Tahoma", 12, "bold"))
        var_label.pack(side="left", padx=5)
        return var_label # Return the label so it can be assigned to an instance variable

    def _create_graph_frame(self):
        """
        สร้างและจัดวางองค์ประกอบ UI สำหรับหน้าจอกราฟ
        """
        self.graph_display_frame = ttk.Frame(self.container_frame, padding=10)
        self.frames["graph"] = self.graph_display_frame
        
        # กำหนด grid สำหรับ graph_display_frame เพื่อจัดการตำแหน่งของปุ่มและกราฟ
        self.graph_display_frame.grid_rowconfigure(1, weight=1) # แถวสำหรับกราฟ ให้ขยายเต็ม
        self.graph_display_frame.grid_columnconfigure(0, weight=1) # คอลัมน์สำหรับกราฟ ให้ขยายเต็ม

        # สร้างปุ่มกลับหน้าหลักและวางไว้ในตำแหน่งบนสุดด้านซ้ายสุด
        self.back_button = ttk.Button(self.graph_display_frame, text="กลับหน้าหลัก", command=lambda: self.show_frame("main"))
        self.back_button.grid(row=0, column=0, sticky="nw", padx=10, pady=10)

        # self.canvas and self.toolbar will be created dynamically in show_graph_frame

    def show_frame(self, frame_name):
        """
        แสดง Frame ที่ระบุและซ่อน Frame อื่นๆ
        """
        for f in self.frames.values():
            f.grid_forget() # ซ่อน Frame ปัจจุบันทั้งหมด
        
        frame = self.frames[frame_name]
        # แสดง Frame ที่ต้องการ และให้ขยายเต็มพื้นที่ของ container_frame
        frame.grid(row=0, column=0, sticky="nsew", in_=self.container_frame)

    def has_internet(self):
        """
        ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต
        """
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=2)
            return True
        except OSError:
            return False

    def find_support_levels(self, hist):
        """
        คำนวณระดับแนวรับ 3 ระดับ
        """
        close_prices = hist['Close']
        
        if close_prices.empty:
            return []

        # หาจุดต่ำสุดในท้องถิ่น (local minima)
        local_minima = close_prices[(close_prices.shift(1) > close_prices) & (close_prices.shift(-1) > close_prices)]

        # เลือก 3 จุดต่ำสุดที่เล็กที่สุด (แนวรับที่แข็งแกร่งที่สุด)
        support_levels = local_minima.nsmallest(3).sort_values(ascending=True)

        final_supports = support_levels.tolist()

        # หากมีแนวรับไม่ถึง 3 ระดับ ให้เพิ่มระดับโดยใช้ค่าต่ำสุดรวม หรือคำนวณจากราคาปัจจุบัน
        if len(final_supports) < 3:
            min_overall = close_prices.min()
            
            # เพิ่มค่าต่ำสุดรวมหากยังไม่มีอยู่ในแนวรับที่พบ
            if not any(abs(lvl - min_overall) < 0.01 for lvl in final_supports):
                final_supports.append(min_overall)
            
            final_supports = sorted(list(set([round(x, 2) for x in final_supports])))

            while len(final_supports) < 3:
                if final_supports:
                    last_known_support = sorted(final_supports)[0] # ใช้แนวรับที่ต่ำที่สุดที่พบ
                    new_lower_support = last_known_support * 0.95 # ลดลง 5% จากแนวรับนั้น
                    
                    # ตรวจสอบไม่ให้ค่าแนวรับที่สร้างขึ้นต่ำกว่า 90% ของค่าต่ำสุดในประวัติ
                    if new_lower_support < close_prices.min() * 0.9:
                        new_lower_support = close_prices.min() * 0.9
                    
                    # เพิ่มเข้าไปถ้ายังไม่มีค่าใกล้เคียง
                    if not any(abs(lvl - new_lower_support) < 0.01 for lvl in final_supports):
                        final_supports.append(new_lower_support)
                    else: # ถ้าค่าที่คำนวณซ้ำกัน ให้หยุด
                        break 
                else: # กรณีไม่พบแนวรับเลย (ข้อมูลน้อยมาก)
                    last_price = close_prices.iloc[-1] if not close_prices.empty else 100
                    final_supports.extend([last_price * 0.98, last_price * 0.95, last_price * 0.92])
                    final_supports = sorted(list(set([round(x, 2) for x in final_supports])))[:3]
                    break
        
        return sorted(list(set([round(x, 2) for x in final_supports])))[:3]

    def find_resistance_levels(self, hist):
        """
        คำนวณระดับแนวต้าน 3 ระดับ
        """
        close_prices = hist['Close']
        
        if close_prices.empty:
            return []

        # หาจุดสูงสุดในท้องถิ่น (local maxima)
        local_maxima = close_prices[(close_prices.shift(1) < close_prices) & (close_prices.shift(-1) < close_prices)]
        
        # เลือก 3 จุดสูงสุดที่ใหญ่ที่สุด (แนวต้านที่แข็งแกร่งที่สุด)
        resistance_levels = local_maxima.nlargest(3).sort_values(ascending=True) 
        
        final_resistances = resistance_levels.tolist()

        # หากมีแนวต้านไม่ถึง 3 ระดับ ให้เพิ่มระดับโดยใช้ค่าสูงสุดรวม หรือคำนวณจากราคาปัจจุบัน
        if len(final_resistances) < 3:
            max_overall = close_prices.max()
            if not any(abs(lvl - max_overall) < 0.01 for lvl in final_resistances):
                final_resistances.append(max_overall)

            final_resistances = sorted(list(set([round(x, 2) for x in final_resistances])))

            while len(final_resistances) < 3:
                if final_resistances:
                    last_known_resistance = sorted(final_resistances, reverse=True)[0] # ใช้แนวต้านที่สูงสุดที่พบ
                    new_higher_resistance = last_known_resistance * 1.05 # เพิ่มขึ้น 5% จากแนวต้านนั้น
                    
                    # ตรวจสอบไม่ให้ค่าแนวต้านที่สร้างขึ้นสูงกว่า 110% ของค่าสูงสุดในประวัติ
                    if new_higher_resistance > close_prices.max() * 1.1:
                           new_higher_resistance = close_prices.max() * 1.1
                    
                    # เพิ่มเข้าไปถ้ายังไม่มีค่าใกล้เคียง
                    if not any(abs(lvl - new_higher_resistance) < 0.01 for lvl in final_resistances):
                        final_resistances.append(new_higher_resistance)
                    else: # ถ้าค่าที่คำนวณซ้ำกัน ให้หยุด
                        break
                else: # กรณีไม่พบแนวต้านเลย (ข้อมูลน้อยมาก)
                    last_price = close_prices.iloc[-1] if not close_prices.empty else 100
                    final_resistances.extend([last_price * 1.02, last_price * 1.05, last_price * 1.08])
                    final_resistances = sorted(list(set([round(x, 2) for x in final_resistances])))[-3:]
                    break
        
        return sorted(list(set([round(x, 2) for x in final_resistances])))[-3:]

    def calculate_rsi(self, series, window=14):
        """
        คำนวณค่า Relative Strength Index (RSI)
        """
        delta = series.diff()
        gain = delta.where(delta > 0, 0)
        loss = -delta.where(delta < 0, 0)

        avg_gain = gain.ewm(com=window-1, adjust=False).mean()
        avg_loss = loss.ewm(com=window-1, adjust=False).mean()

        rs = avg_gain / avg_loss.replace(0, 1e-10) # ป้องกันหารด้วยศูนย์
        rsi = 100 - (100 / (1 + rs))
        return rsi

    def save_graph_pdf(self):
        """
        บันทึกกราฟเป็นไฟล์ PDF
        """
        if self.current_hist_data is None:
            messagebox.showwarning("คำเตือน", "กรุณาคำนวณข้อมูลหุ้นก่อนบันทึกกราฟ")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=f"{self.symbol_var.get().upper()}_stock_analysis.pdf"
        )
        if file_path:
            try:
                # สร้าง Figure ชั่วคราวเพื่อบันทึก
                temp_figure = self._create_graph_figure(self.current_hist_data, self.current_support_levels, self.current_resistance_levels, self.current_symbol_info)
                temp_figure.savefig(file_path, format='pdf', bbox_inches='tight')
                plt.close(temp_figure) # ปิด Figure ชั่วคราวเพื่อไม่ให้ค้างในหน่วยความจำ
                messagebox.showinfo("บันทึกสำเร็จ", f"บันทึกไฟล์ PDF: {file_path}")
            except Exception as e:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่สามารถบันทึกไฟล์ PDF ได้\n{e}")

    def _load_watchlist(self):
        """
        โหลดรายการหุ้นโปรดจากไฟล์
        """
        try:
            with open(self.watchlist_file, "r") as f:
                return [line.strip().upper() for line in f if line.strip()]
        except FileNotFoundError:
            return []

    def _save_watchlist(self):
        """
        บันทึกรายการหุ้นโปรดลงไฟล์
        """
        with open(self.watchlist_file, "w") as f:
            for symbol in self.watchlist:
                f.write(symbol + "\n")

    def _update_watchlist_listbox(self):
        """
        อัปเดต Listbox แสดงรายการหุ้นโปรด
        """
        # Ensure watchlist_listbox is initialized before attempting to use it
        if self.watchlist_listbox:
            self.watchlist_listbox.delete(0, tk.END)
            for symbol in self.watchlist:
                self.watchlist_listbox.insert(tk.END, symbol)

    def _add_to_watchlist(self):
        """
        เพิ่มหุ้นเข้าสู่รายการโปรด
        """
        symbol = self.symbol_var.get().strip().upper()
        if not symbol:
            messagebox.showwarning("คำเตือน", "กรุณากรอกรหัสหุ้นที่จะเพิ่ม")
            return
        if not symbol.isalnum():
            messagebox.showwarning("คำเตือน", "รหัสหุ้นไม่ถูกต้อง (ตัวอักษรหรือตัวเลขเท่านั้น)")
            return

        if symbol not in self.watchlist:
            self.watchlist.append(symbol)
            self._update_watchlist_listbox()
            self._save_watchlist()
            messagebox.showinfo("สำเร็จ", f"เพิ่ม {symbol} เข้าสู่รายการโปรดแล้ว")
        else:
            messagebox.showinfo("แจ้งเตือน", f"{symbol} มีอยู่ในรายการโปรดแล้ว")

    def _remove_from_watchlist(self):
        """
        ลบหุ้นออกจากรายการโปรด
        """
        selected_index = self.watchlist_listbox.curselection()
        if selected_index:
            symbol_to_remove = self.watchlist_listbox.get(selected_index[0])
            self.watchlist.pop(selected_index[0])
            self._update_watchlist_listbox()
            self._save_watchlist()
            messagebox.showinfo("สำเร็จ", f"ลบ {symbol_to_remove} ออกจากรายการโปรดแล้ว")
        else:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกหุ้นที่ต้องการลบ")
            
    def _on_watchlist_select(self, event):
        """
        เมื่อเลือกหุ้นจาก watchlist จะนำรหัสหุ้นไปใส่ในช่องกรอกและคำนวณทันที
        """
        selected_index = self.watchlist_listbox.curselection()
        if selected_index:
            symbol = self.watchlist_listbox.get(selected_index[0])
            self.symbol_var.set(symbol)
            self.calculate_support()

    def _on_closing(self):
        """
        จัดการการบันทึก watchlist ก่อนปิดโปรแกรม
        """
        self._save_watchlist()
        self.root.destroy()

    def _create_graph_figure(self, hist, supports, resistances, info):
        """
        สร้าง Figure Matplotlib สำหรับแสดงกราฟราคา, RSI และ Volume
        """
        # สร้าง subplot 3 แถว โดยแถวบนสุด (ราคา) มีขนาดใหญ่กว่า
        figure = plt.Figure(figsize=(10, 7), dpi=100, facecolor="white")
        gs = figure.add_gridspec(3, 1, height_ratios=[3, 1, 1], hspace=0) # hspace=0 เพื่อให้กราฟติดกัน
        ax_price = figure.add_subplot(gs[0, 0])
        ax_rsi = figure.add_subplot(gs[1, 0], sharex=ax_price) # แชร์แกน X กับ ax_price
        ax_volume = figure.add_subplot(gs[2, 0], sharex=ax_price) # แชร์แกน X กับ ax_price

        # พล็อตกราฟราคาปิด
        ax_price.plot(hist.index, hist['Close'], label="ราคาปิด", color="#0077cc")
        
        # พล็อตกราฟ Moving Average 50 วัน
        hist['MA50'] = hist['Close'].rolling(window=50).mean()
        ax_price.plot(hist.index, hist['MA50'], label="MA 50", color="orange")
        # พล็อตกราฟ Moving Average 200 วัน (ถ้ามีข้อมูลพอ)
        if len(hist) >= 200:
            hist['MA200'] = hist['Close'].rolling(window=200).mean()
            ax_price.plot(hist.index, hist['MA200'], label="MA 200", color="green")
        
        # พล็อตเส้นแนวรับ
        for i, level in enumerate(supports):
            label_text = f"แนวรับต่ำสุด ({level:.2f})" if i == 0 else \
                         f"แนวรับรอง ({level:.2f})" if i == 1 else \
                         f"แนวรับสูงสุด ({level:.2f})"
            ax_price.axhline(level, color="red", linestyle="--", linewidth=1, label=label_text)
        
        # พล็อตเส้นแนวต้าน
        for i, level in enumerate(resistances):
            label_text = f"แนวต้านต่ำสุด ({level:.2f})" if i == 0 else \
                         f"แนวต้านรอง ({level:.2f})" if i == 1 else \
                         f"แนวต้านสูงสุด ({level:.2f})"
            ax_price.axhline(level, color="purple", linestyle="--", linewidth=1, label=label_text)

        ax_price.set_title(f"กราฟราคา {info.get('longName', info.get('symbol', 'N/A'))} และตัวชี้วัด", loc='left')
        ax_price.set_ylabel("ราคา (USD)")
        ax_price.legend(loc="upper left", fontsize=8)
        ax_price.grid(True)
        plt.setp(ax_price.get_xticklabels(), visible=False) # ซ่อน label แกน X เพื่อความเรียบร้อย

        # พล็อตกราฟ RSI
        hist['RSI'] = self.calculate_rsi(hist['Close'])
        ax_rsi.plot(hist.index, hist['RSI'], label="RSI (14)", color="blue")
        ax_rsi.axhline(70, color='red', linestyle=':', linewidth=0.8, label='Overbought (70)')
        ax_rsi.axhline(30, color='green', linestyle=':', linewidth=0.8, label='Oversold (30)')
        ax_rsi.set_ylabel("RSI")
        ax_rsi.set_ylim(0, 100) # กำหนดช่วง RSI
        ax_rsi.legend(loc="upper left", fontsize=8)
        ax_rsi.grid(True)
        plt.setp(ax_rsi.get_xticklabels(), visible=False) # ซ่อน label แกน X เพื่อความเรียบร้อย

        # พล็อตกราฟ Volume
        ax_volume.bar(hist.index, hist['Volume'], color='gray', alpha=0.7, label="ปริมาณการซื้อขาย")
        ax_volume.set_ylabel("Volume")
        ax_volume.set_xlabel("วันที่")
        ax_volume.ticklabel_format(style='plain', axis='y') # แสดงเลขแบบปกติ
        ax_volume.legend(loc="upper left", fontsize=8)
        ax_volume.grid(True)

        figure.tight_layout(rect=[0, 0, 1, 0.96]) # ปรับ layout ให้สวยงาม
        return figure

    def show_graph_frame(self):
        """
        แสดงหน้าจอกราฟพร้อมข้อมูลที่คำนวณไว้แล้ว
        """
        if self.current_hist_data is None:
            messagebox.showwarning("คำเตือน", "กรุณาคำนวณข้อมูลหุ้นก่อนดูกราฟ")
            return

        # ล้าง widgets เก่าใน graph_display_frame ที่เกี่ยวข้องกับกราฟและ toolbar
        for widget in self.graph_display_frame.winfo_children():
            if widget != self.back_button: # ไม่ต้องทำลาย back_button
                widget.destroy()
        
        # สร้าง Figure ใหม่และ Canvas ใหม่ทุกครั้ง
        self.figure = self._create_graph_figure(self.current_hist_data, self.current_support_levels, self.current_resistance_levels, self.current_symbol_info)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.graph_display_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        # วาง canvas ในแถวที่ 1, คอลัมน์ 0 และให้ขยายเต็มพื้นที่ที่เหลือ
        self.canvas_widget.grid(row=1, column=0, sticky="nsew", columnspan=2) 

        # สร้าง toolbar ใหม่
        self.toolbar_frame = ttk.Frame(self.graph_display_frame)
        self.toolbar_frame.grid(row=0, column=1, sticky="ew") # วางในแถวเดียวกับปุ่ม แต่คอลัมน์ขวา
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.toolbar_frame)
        self.toolbar.update()
        
        self.canvas.draw() # วาดกราฟ
        self.show_frame("graph") # สลับไปแสดงหน้าจอกราฟ

    def calculate_support(self):
        """
        ดึงข้อมูลหุ้น คำนวณแนวรับ แนวต้าน และอัปเดต UI
        """
        symbol = self.symbol_var.get().strip().upper()
        period = self.period_var.get()

        if not symbol:
            messagebox.showwarning("คำเตือน", "กรุณากรอกรหัสหุ้น")
            return

        if not self.has_internet():
            messagebox.showerror("ข้อผิดพลาด", "ไม่มีการเชื่อมต่ออินเทอร์เน็ต")
            return

        if not symbol.isalnum():
            messagebox.showerror("ข้อผิดพลาด", "รหัสหุ้นไม่ถูกต้อง: ควรเป็นตัวอักษรหรือตัวเลขเท่านั้น")
            return
        
        # เพิ่มรหัสหุ้นลงในประวัติการค้นหา
        if symbol and symbol not in self.search_history:
            self.search_history.insert(0, symbol)
            self.search_history = self.search_history[:10] # เก็บ 10 รายการล่าสุด
            self.symbol_entry['values'] = self.search_history

        # อัปเดตสถานะ UI และปิดการใช้งานปุ่มชั่วคราว
        self.status_label.config(text="กำลังโหลดข้อมูล...")
        self.calc_button.config(state="disabled")
        self.view_graph_button.config(state="disabled")
        self.save_pdf_button.config(state="disabled")
        self.root.update_idletasks() # บังคับให้อัปเดต UI ทันที

        try:
            ticker = yf.Ticker(symbol)
            info = ticker.info # ดึงข้อมูลบริษัท

            if not info or "regularMarketPrice" not in info or info["regularMarketPrice"] is None:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่พบข้อมูลหุ้นที่ถูกต้องสำหรับ: {symbol} หรือไม่มีราคาปัจจุบัน")
                self.status_label.config(text="")
                return

            hist = ticker.history(period=period) # ดึงข้อมูลประวัติราคา
            if hist.empty:
                messagebox.showerror("ข้อผิดพลาด", f"ไม่มีข้อมูลประวัติหุ้นสำหรับ {symbol} ในช่วงเวลาที่เลือก")
                self.status_label.config(text="")
                return

            # คำนวณแนวรับและแนวต้าน
            supports = self.find_support_levels(hist)
            resistances = self.find_resistance_levels(hist)
            
            # อัปเดตข้อมูลบริษัทใน UI
            self.company_label.config(text=info.get("longName", symbol))
            # Access the labels directly after they have been assigned in _create_main_frame
            self.sector_label.config(text=info.get("sector", "N/A"))
            self.industry_label.config(text=info.get("industry", "N/A"))
            
            market_cap = info.get("marketCap")
            self.marketcap_label.config(text=f"{market_cap:,.0f} USD" if market_cap else "N/A")

            trailing_pe = info.get("trailingPE")
            self.pe_label.config(text=f"{trailing_pe:.2f}" if trailing_pe else "N/A")

            dividend_yield = info.get("dividendYield")
            self.dividend_label.config(text=f"{dividend_yield*100:.2f}%" if dividend_yield else "N/A")
            
            # อัปเดตราคาล่าสุดและแนวรับ แนวต้าน
            last_close = hist['Close'].iloc[-1]
            self.price_label.config(text=f"ราคาปิดล่าสุด: {last_close:.2f} USD")
            
            for i, level in enumerate(supports):
                self.support_labels[i].config(text=f"{level:.2f} USD")
            for i, level in enumerate(resistances):
                self.resistance_labels[i].config(text=f"{level:.2f} USD")

            self.status_label.config(text=f"แสดงข้อมูล {len(hist)} วันย้อนหลัง")

            # เก็บข้อมูลที่คำนวณได้ เพื่อใช้ในการแสดงกราฟหรือ PDF
            self.current_hist_data = hist
            self.current_support_levels = supports
            self.current_resistance_levels = resistances
            self.current_symbol_info = info

            # เปิดใช้งานปุ่มดูกราฟและบันทึก PDF
            self.view_graph_button.config(state="normal")
            self.save_pdf_button.config(state="normal")

        except Exception as e:
            messagebox.showerror("ข้อผิดพลาด", f"เกิดข้อผิดพลาดในการดึงข้อมูลหรือคำนวณ\n{e}")
            self.status_label.config(text="เกิดข้อผิดพลาดในการโหลดข้อมูล")
            # รีเซ็ตข้อมูลที่เก็บไว้
            self.current_hist_data = None
            self.current_symbol_info = None
            self.view_graph_button.config(state="disabled")
            self.save_pdf_button.config(state="disabled")
        finally:
            self.calc_button.config(state="normal")
            # อัปเดตสถานะสุดท้าย
            if "ข้อผิดพลาด" not in self.status_label.cget("text"):
                self.status_label.config(text=f"ข้อมูลพร้อมสำหรับ {symbol}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StockSupportCalculator(root)
    root.mainloop()