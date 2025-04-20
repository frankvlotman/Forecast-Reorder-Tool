import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from PIL import Image
import tkinter.font as tkFont

# --- Tooltip helper ---
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.showtip)
        widget.bind("<Leave>", self.hidetip)

    def showtip(self, event):
        if self.tipwindow or not self.text:
            return
        x = event.x_root + 20
        y = event.y_root + 10
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw,
            text=self.text,
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", "8", "normal")
        )
        label.pack(ipadx=1)

    def hidetip(self, event):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

# --- Create blank .ico if missing ---
def create_blank_ico(path):
    size = (16, 16)
    img = Image.new("RGBA", size, (255, 255, 255, 0))
    img.save(path, format="ICO")

ICON_PATH = r"C:\Users\Frank\Desktop\blank.ico"
create_blank_ico(ICON_PATH)

class ReorderCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Future Purchase Reorder Quantities")
        self.geometry("1000x700")
        self.iconbitmap(ICON_PATH)

        # Modern styling
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('Segoe UI', 10), padding=5)
        style.configure('TButton', font=('Segoe UI', 10), padding=5)
        style.configure('TEntry', font=('Segoe UI', 10))
        style.map('TButton', background=[('active', '#87CEFA')])

        # Layout
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

        # Build sections
        self.create_header()
        self.create_input_frame()
        self.create_button_frame()

        # Placeholders
        self.data_frame = None
        self.output_frame = None

    def create_header(self):
        hdr = ttk.Label(
            self,
            text="Reorder Quantity Calculator",
            font=("Segoe UI", 16, "bold")
        )
        hdr.grid(row=0, column=0, pady=(10, 0))

    def create_input_frame(self):
        f = ttk.LabelFrame(self, text="Input Parameters", padding=10)
        f.grid(row=1, column=0, sticky='ew', padx=20, pady=10)
        for c in range(4):
            f.columnconfigure(c, weight=1)

        # Months Ahead
        ttk.Label(f, text="Months Ahead:").grid(row=0, column=0, sticky='e')
        self.months_ahead = ttk.Entry(f, width=12)
        self.months_ahead.grid(row=0, column=1, sticky='w')
        self.months_ahead.bind("<Return>", self.focus_next_widget)
        ToolTip(self.months_ahead, "How many months into the future to forecast.")

        # Starting Month
        ttk.Label(f, text="Starting Month:").grid(row=0, column=2, sticky='e')
        self.starting_month = ttk.Combobox(
            f,
            values=[f"{m} 2024" for m in
                    ['Jan','Feb','Mar','Apr','May','Jun',
                     'Jul','Aug','Sep','Oct','Nov','Dec']],
            width=12
        )
        self.starting_month.grid(row=0, column=3, sticky='w')
        self.starting_month.bind("<Return>", self.focus_next_widget)
        ToolTip(self.starting_month, "Select the first month for your forecast.")

        # Opening Stock
        ttk.Label(f, text="Opening Stock:").grid(row=1, column=0, sticky='e')
        self.opening_stock = ttk.Entry(f, width=12)
        self.opening_stock.grid(row=1, column=1, sticky='w')
        self.opening_stock.bind("<Return>", self.focus_next_widget)
        ToolTip(self.opening_stock, "Current on-hand inventory at the start.")

        # Target Months Stock
        ttk.Label(f, text="Target Months Stock:").grid(row=1, column=2, sticky='e')
        self.target_months_stock = ttk.Entry(f, width=12)
        self.target_months_stock.grid(row=1, column=3, sticky='w')
        self.target_months_stock.bind("<Return>", self.focus_next_widget)
        ToolTip(
            self.target_months_stock,
            "Number of months of forecast sales to hold as safety stock."
        )

    def create_button_frame(self):
        f = ttk.Frame(self)
        f.grid(row=2, column=0, pady=10)
        f.columnconfigure((0, 1, 2), weight=1)

        ttk.Button(f, text="Generate Table", command=self.generate_table).grid(row=0, column=0, padx=5)
        ttk.Button(f, text="Paste Forecast", command=lambda: self.paste_from_clipboard(0)).grid(row=0, column=1, padx=5)
        ttk.Button(f, text="Paste Qty to Order", command=lambda: self.paste_from_clipboard(1)).grid(row=0, column=2, padx=5)

    def generate_table(self):
        # Validate
        try:
            months = int(self.months_ahead.get())
            start = datetime.strptime(self.starting_month.get(), "%b %Y")
            opening = int(self.opening_stock.get())
            self.target_mths = float(self.target_months_stock.get())
        except Exception:
            messagebox.showerror("Input Error", "Please check your inputs.")
            return

        # Month labels
        self.month_names = [
            (start + relativedelta(months=i)).strftime("%b %Y")
            for i in range(months)
        ]

        # Cleanup old frames
        if self.data_frame:
            self.data_frame.destroy()
        if self.output_frame:
            self.output_frame.destroy()

        # Input grid
        self.data_frame = ttk.Frame(self)
        self.data_frame.grid(row=3, column=0, sticky='nsew', padx=20)

        ttk.Label(self.data_frame, text="Parameter").grid(row=0, column=0, padx=5)
        for c, m in enumerate(self.month_names, start=1):
            ttk.Label(self.data_frame, text=m).grid(row=0, column=c, padx=5)

        labels = ["Forecast Sales", "Qty to Order", "Opening Stock Balance"]
        self.entries = []
        for r, txt in enumerate(labels, start=1):
            ttk.Label(self.data_frame, text=txt).grid(row=r, column=0, sticky='e', padx=5, pady=2)
            row_e = []
            for c in range(months):
                e = ttk.Entry(self.data_frame, width=10)
                e.grid(row=r, column=c+1, padx=2, pady=2)
                e.bind("<Return>", self.focus_next_widget)
                row_e.append(e)
            self.entries.append(row_e)

        # Prefill opening
        self.entries[2][0].insert(0, opening)
        self.entries[2][0].state(['readonly'])

        # Control buttons
        ctrl = ttk.Frame(self)
        ctrl.grid(row=4, column=0, pady=10)
        ttk.Button(ctrl, text="Calculate Results", command=self.calculate_closing_stock).grid(row=0, column=0, padx=5)
        ttk.Button(ctrl, text="Download XLSX", command=self.download_to_xlsx).grid(row=0, column=1, padx=5)
        ttk.Button(ctrl, text="Scenario Analysis", command=self.open_scenario_window).grid(row=0, column=2, padx=5)

        # Output grid
        self.setup_output_table()

    def setup_output_table(self):
        self.output_frame = ttk.Frame(self)
        self.output_frame.grid(row=5, column=0, padx=20, pady=10, sticky='nsew')

        cols = ['Parameter'] + self.month_names
        self.tree = ttk.Treeview(self.output_frame, columns=cols, show='headings')
        style = ttk.Style()
        style.configure("Treeview.Heading", background="light blue")
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100, anchor='center')

        params = [
            "Forecast Sales","Qty to Order","Opening Stock Balance",
            "Closing Stock Balance","Months Stock","Target Months Stock",
            "Suggest Qty to Order"
        ]
        tags = ['evenrow','oddrow']
        for i, p in enumerate(params):
            vals = [p] + (['0'] if p=="Opening Stock Balance" else [""]*len(self.month_names))
            self.tree.insert('', 'end', values=vals, tags=(tags[i%2],))

        self.tree.tag_configure('evenrow', background='white')
        self.tree.tag_configure('oddrow', background='light grey')

        sb = ttk.Scrollbar(self.output_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        sb.grid(row=0, column=1, sticky='ns')

        self.autofit_columns()

    def calculate_closing_stock(self):
        opening = int(self.opening_stock.get())
        curr = opening
        rows = self.tree.get_children()
        for i, m in enumerate(self.month_names):
            f = int(self.entries[0][i].get() or 0)
            o = int(self.entries[1][i].get() or 0)
            close = curr - f + o
            ms = round(close / f, 3) if f else 0
            sugg = round((self.target_mths - ms) * f)

            self.tree.set(rows[0], m, f)
            self.tree.set(rows[1], m, o)
            # Opening Stock Balance column retains its first value
            self.tree.set(rows[2], m, opening if i==0 else self.tree.set(rows[2], m))
            self.tree.set(rows[3], m, close)
            self.tree.set(rows[4], m, f"{ms:.3f}")
            self.tree.set(rows[5], m, str(self.target_mths))
            self.tree.set(rows[6], m, sugg)

            curr = close

    def open_scenario_window(self):
        opening   = int(self.opening_stock.get())
        target    = self.target_mths
        months    = self.month_names
        forecasts = [int(e.get() or 0) for e in self.entries[0]]
        on_orders = [int(e.get() or 0) for e in self.entries[1]]

        # Scenario calculations
        scen_on, scen_cl, scen_ms, scen_sugg = [], [], [], []
        tmp = opening
        for f in forecasts:
            req = round((target - ((tmp - f)/f if f else 0)) * f)
            scen_on.append(req)
            c = tmp - f + req
            ms = round(c/f, 3) if f else 0
            s = round((target - ms) * f)
            scen_cl.append(c); scen_ms.append(ms); scen_sugg.append(s)
            tmp = c
        scen_open = [opening] + scen_cl[:-1]
        self.scenario_data = {'opening':opening,'target':target,'months':months}

        win = tk.Toplevel(self)
        win.title("Scenario Analysis")
        rf = ttk.LabelFrame(win, text="Scenario (editable)", padding=10)
        rf.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        cols = ['Parameter'] + months
        scen_tree = ttk.Treeview(rf, columns=cols, show='headings')
        for c in cols:
            scen_tree.heading(c, text=c)
            scen_tree.column(c, width=80, anchor='center')
        scen_tree.grid(row=0, column=0, sticky='nsew')
        sb = ttk.Scrollbar(rf, orient='vertical', command=scen_tree.yview)
        scen_tree.configure(yscroll=sb.set)
        sb.grid(row=0, column=1, sticky='ns')

        params = ["Opening Stock","Forecast Sales","Qty to Order",
                  "Closing Stock","Months Stock","Suggest Qty to Order"]
        self.scen_items = {}
        for p in params:
            if p=="Opening Stock":
                cv, sv = scen_open, scen_open
            elif p=="Forecast Sales":
                cv, sv = forecasts, forecasts
            elif p=="Qty to Order":
                cv, sv = on_orders, scen_on
            elif p=="Closing Stock":
                cv, sv = scen_cl, scen_cl
            elif p=="Months Stock":
                cv, sv = scen_ms, scen_ms
            else:
                cv, sv = scen_sugg, scen_sugg
            iid = scen_tree.insert('', 'end', values=[p] + [str(x) for x in sv])
            self.scen_items[p] = iid

        scen_tree.bind('<Double-1>', self._on_scenario_double_click)
        self.scen_tree = scen_tree

    def _on_scenario_double_click(self, event):
        tree  = event.widget
        rowid = tree.identify_row(event.y)
        if rowid not in (
            self.scen_items['Forecast Sales'],
            self.scen_items['Qty to Order']
        ):
            return
        col = tree.identify_column(event.x)
        idx = int(col.lstrip('#')) - 2
        if idx < 0 or idx >= len(self.month_names):
            return
        col_name = self.month_names[idx]
        x, y, w, h = tree.bbox(rowid, col)
        entry = tk.Entry(tree)
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, tree.set(rowid, col_name))
        entry.focus()
        self._edit_info = (entry, rowid, col_name)
        entry.bind('<Return>', self._commit_edit)
        entry.bind('<FocusOut>', self._commit_edit)

    def _commit_edit(self, event):
        entry, rowid, col = self._edit_info
        new = entry.get()
        entry.destroy()
        self.scen_tree.set(rowid, column=col, value=new)
        self._update_scenario()
        self._edit_info = None

    def _update_scenario(self):
        d      = self.scenario_data
        months = d['months']; opening = d['opening']; target = d['target']
        fcasts = [int(self.scen_tree.set(self.scen_items['Forecast Sales'], m) or 0)
                  for m in months]
        onords = [int(self.scen_tree.set(self.scen_items['Qty to Order'], m) or 0)
                  for m in months]
        curr = opening
        cl_list, ms_list, sugg_list = [], [], []
        for f, o in zip(fcasts, onords):
            c = curr - f + o
            ms = round(c/f, 3) if f else 0
            s = round((target-ms)*f)
            cl_list.append(c); ms_list.append(ms); sugg_list.append(s)
            curr = c
        scen_open = [opening] + cl_list[:-1]
        for i, m in enumerate(months):
            self.scen_tree.set(self.scen_items['Opening Stock'], m, str(scen_open[i]))
            self.scen_tree.set(self.scen_items['Closing Stock'], m, str(cl_list[i]))
            self.scen_tree.set(self.scen_items['Months Stock'], m, f"{ms_list[i]:.3f}")
            self.scen_tree.set(self.scen_items['Suggest Qty to Order'], m, str(sugg_list[i]))

    def autofit_columns(self):
        font = tkFont.Font()
        for c in self.tree['columns']:
            maxw = font.measure(c)
            for iid in self.tree.get_children():
                w = font.measure(str(self.tree.set(iid, c)))
                if w > maxw:
                    maxw = w
            self.tree.column(c, width=maxw+10)

    def download_to_xlsx(self):
        wb = openpyxl.Workbook()
        ws = wb.active; ws.title = "Reorder Quantities"
        ws.append(['Parameter'] + self.month_names)
        for iid in self.tree.get_children():
            ws.append(self.tree.item(iid)['values'])
        path = r"C:\Users\Frank\Desktop\Reorder_Quantities1.xlsx"
        wb.save(path)
        messagebox.showinfo("Saved", f"Saved to {path}")

    def paste_from_clipboard(self, row):
        try:
            vals = self.clipboard_get().strip().split()
            if len(vals) != len(self.month_names):
                raise ValueError
            for i, v in enumerate(vals):
                e = self.entries[row][i]
                e.delete(0, 'end')
                e.insert(0, v)
        except:
            messagebox.showerror("Error", "Clipboard data mismatch or error")

    def focus_next_widget(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

if __name__ == "__main__":
    ReorderCalculator().mainloop()
