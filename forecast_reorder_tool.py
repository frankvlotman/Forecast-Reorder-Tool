import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl
from PIL import Image
import tkinter.font as tkFont  # Import the font module

# Define the path for the blank icon
icon_path = 'C:\\Users\\Frank\\Desktop\\blank.ico'

# Create a blank (transparent) ICO file if it doesn't exist
def create_blank_ico(path):
    size = (16, 16)  # Size of the icon
    image = Image.new("RGBA", size, (255, 255, 255, 0))  # Transparent image
    image.save(path, format="ICO")

# Create the blank ICO file
create_blank_ico(icon_path)

class ReorderCalculator(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Future Purchase Reorder Quantities")
        self.geometry("1327x900")  # Adjusted size for better visibility
        self.iconbitmap(icon_path)  # Set the window icon
        
        self.create_widgets()

    def create_widgets(self):
        # Input fields for number of months and starting month
        tk.Label(self, text="Number of months ahead:").grid(row=0, column=0, padx=10, pady=10)
        self.months_ahead = tk.Entry(self, width=10)
        self.months_ahead.grid(row=0, column=1, padx=10, pady=10)
        self.months_ahead.bind("<Return>", self.focus_next_widget)

        tk.Label(self, text="Starting month (e.g., Jan 2024):").grid(row=1, column=0, padx=10, pady=10)
        
        # Replace tk.Entry with ttk.Combobox with adjusted width
        self.starting_month = ttk.Combobox(self, values=[
            'Jan 2024', 'Feb 2024', 'Mar 2024', 'Apr 2024', 'May 2024', 'Jun 2024', 'Jul 2024', 'Aug 2024', 'Sep 2024', 'Oct 2024', 'Nov 2024', 'Dec 2024'
        ], width=12)
        self.starting_month.grid(row=1, column=1, padx=10, pady=10)
        self.starting_month.bind("<Return>", self.focus_next_widget)

        tk.Label(self, text="Opening Stock Balance:").grid(row=2, column=0, padx=10, pady=10)
        self.opening_stock = tk.Entry(self, width=10)
        self.opening_stock.grid(row=2, column=1, padx=10, pady=10)
        self.opening_stock.bind("<Return>", self.focus_next_widget)

        tk.Label(self, text="Target Months Stock Holding:").grid(row=3, column=0, padx=10, pady=10)
        self.target_months_stock = tk.Entry(self, width=10)
        self.target_months_stock.grid(row=3, column=1, padx=10, pady=10)
        self.target_months_stock.bind("<Return>", self.focus_next_widget)

        self.target_months_stock_note = tk.Label(self, text="Input '0' if only need to fulfil Forecast Sales without need for safety stock")
        self.target_months_stock_note.grid(row=3, column=2, padx=10, pady=10, sticky='w')

        self.submit_button = tk.Button(self, text="Enter", command=self.generate_table)
        self.submit_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
        self.submit_button.bind("<Return>", lambda event: self.generate_table())

        # Add buttons for pasting clipboard data for each row
        self.paste_forecast_button = tk.Button(self, text="Paste Forecast Sales", command=lambda: self.paste_from_clipboard(0))
        self.paste_forecast_button.grid(row=4, column=2, columnspan=2, padx=10, pady=10)
        
        self.paste_order_button = tk.Button(self, text="Paste On Order Qty", command=lambda: self.paste_from_clipboard(1))
        self.paste_order_button.grid(row=4, column=4, columnspan=2, padx=10, pady=10)

    def generate_table(self, event=None):
        months_ahead = int(self.months_ahead.get())
        starting_month = self.starting_month.get()
        opening_stock = int(self.opening_stock.get())
        
        start_date = datetime.strptime(starting_month, "%b %Y")
        
        # Create a frame for the inputs and labels
        self.input_frame = ttk.Frame(self)
        self.input_frame.grid(row=5, column=0, columnspan=6, padx=10, pady=10, sticky='nsew')

        self.month_names = []

        # Create column headers for months
        for i in range(months_ahead):
            month = (start_date + relativedelta(months=i)).strftime("%b %Y")
            self.month_names.append(month)
            tk.Label(self.input_frame, text=month).grid(row=0, column=i+1, padx=5, pady=5)

        # Row headers for input types
        input_labels = ["Forecast Sales", "On Order Qty", "Opening Stock Balance"]
        for row, label in enumerate(input_labels, start=1):
            tk.Label(self.input_frame, text=label).grid(row=row, column=0, padx=5, pady=5, sticky='e')

        # Populate the table with input fields
        self.entries = [[None for _ in range(months_ahead)] for _ in range(3)]
        for row in range(3):
            for col in range(months_ahead):
                entry = tk.Entry(self.input_frame, width=8)
                entry.grid(row=row+1, column=col+1, padx=5, pady=5)
                entry.bind("<Return>", self.focus_next_widget)
                self.entries[row][col] = entry
        
        # Prepopulate the opening stock balance for the first month
        self.entries[2][0].insert(0, opening_stock)
        self.entries[2][0].configure(state='readonly')

        self.calculate_button = tk.Button(self, text="Calculate Results", command=self.calculate_closing_stock)
        self.calculate_button.grid(row=6+months_ahead, column=0, columnspan=2, padx=10, pady=10)
        self.calculate_button.bind("<Return>", lambda event: self.calculate_closing_stock())
        
        self.download_button = tk.Button(self, text="Download as XLSX", command=self.download_to_xlsx)
        self.download_button.grid(row=6+months_ahead, column=2, columnspan=2, padx=10, pady=10)
        self.download_button.bind("<Return>", lambda event: self.download_to_xlsx())

        # Create a frame for the output table
        self.output_frame = ttk.Frame(self)
        self.output_frame.grid(row=7+months_ahead, column=0, columnspan=6, padx=10, pady=10, sticky='nsew')

        # Create the TreeView for the output table with alternating row colors
        style = ttk.Style()
        style.configure("Treeview.Heading", background="light blue", font=('Helvetica', 10, 'bold'))

        columns = ['Parameter'] + self.month_names
        self.tree = ttk.Treeview(self.output_frame, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col, anchor='center')
            self.tree.column(col, anchor='center', width=100)
        
        self.tree.tag_configure('oddrow', background='light grey')
        self.tree.tag_configure('evenrow', background='white')
        
        self.tree.grid(row=0, column=0, sticky='nsew')

        self.scrollbar = ttk.Scrollbar(self.output_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.scrollbar.grid(row=0, column=1, sticky='ns')

        # Insert the initial rows for the output table with tags for alternating colors
        parameters = ["Forecast Sales", "On Order Qty", "Opening Stock Bal", "Closing Stock Bal", "Months Stock", "Target Months Stock ", "Suggest Qty to Order"]
        for i, param in enumerate(parameters):
            values = [param] + (["0"] if param == "Opening Stock Balance" else [""] * months_ahead)
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.tree.insert("", "end", values=values, tags=(tag,))

        # Autofit columns to the content
        self.autofit_columns()

        # Add the explanatory text and Copy button at the bottom center
        explanatory_text = (
            "Opening Stock Bal = The previous months Closing Stock Bal\n"
            "Closing Stock Bal = Opening Stock - Forecast Sales + On Order Qty\n"
            "Months Stock = Closing Stock Bal / Forecast Sales\n"
            "Suggested Qty to Order = MAX(0, (Target Months Stock - Months Stock) * Forecast Sales)"
        )
        self.additional_explanatory_label = tk.Label(self, text=explanatory_text, justify='left')
        self.additional_explanatory_label.grid(row=9+months_ahead, column=0, columnspan=5, padx=10, pady=10, sticky='w')

        # Add the Copy button next to the explanatory text
        self.copy_button = tk.Button(self, text="Copy Explanation", command=self.copy_explanatory_text)
        self.copy_button.grid(row=9+months_ahead, column=5, padx=10, pady=10, sticky='w')

    def copy_explanatory_text(self):
        explanatory_text = (
            "Opening Stock Bal = The previous months Closing Stock Bal\n"
            "Closing Stock Bal = Opening Stock - Forecast Sales + On Order Qty\n"
            "Months Stock = Closing Stock Bal / Forecast Sales\n"
            "Suggested Qty to Order = MAX(0, (Target Months Stock - Months Stock) * Forecast Sales)"
        )
        self.clipboard_clear()
        self.clipboard_append(explanatory_text)
        self.update()  # Needed to ensure the clipboard is updated immediately
        messagebox.showinfo("Copied", "Explanatory text copied to clipboard!")

    def calculate_closing_stock(self, event=None):
        opening_stock = int(self.opening_stock.get())
        target_months_stock = float(self.target_months_stock.get())
        current_stock = opening_stock

        for col in range(len(self.month_names)):
            forecast = int(self.entries[0][col].get() or 0)
            supplier_qty = int(self.entries[1][col].get() or 0)
            closing_stock_value = current_stock - forecast + supplier_qty
            current_stock = closing_stock_value
            months_stock = round(closing_stock_value / forecast, 3) if forecast != 0 else 0  # Updated to 3 decimal places
            suggested_qty_to_order = max(0, round((target_months_stock - months_stock) * forecast))  # Using round

            self.tree.set(self.tree.get_children()[0], column=self.month_names[col], value=str(forecast))
            self.tree.set(self.tree.get_children()[1], column=self.month_names[col], value=str(supplier_qty))
            self.tree.set(self.tree.get_children()[2], column=self.month_names[col], value=str(opening_stock if col == 0 else self.entries[2][col].get()))
            self.tree.set(self.tree.get_children()[3], column=self.month_names[col], value=str(closing_stock_value))
            self.tree.set(self.tree.get_children()[4], column=self.month_names[col], value=f"{months_stock:.3f}")  # Updated to show 3 decimal places
            self.tree.set(self.tree.get_children()[5], column=self.month_names[col], value=str(target_months_stock))
            self.tree.set(self.tree.get_children()[6], column=self.month_names[col], value=str(suggested_qty_to_order))

            if col < len(self.entries[2]) - 1:
                next_opening_stock_entry = self.entries[2][col + 1]
                next_opening_stock_entry.configure(state='normal')
                next_opening_stock_entry.delete(0, tk.END)
                next_opening_stock_entry.insert(0, closing_stock_value)
                next_opening_stock_entry.configure(state='readonly')

    def autofit_columns(self):
        for col in self.tree["columns"]:
            max_width = tkFont.Font().measure(col)  # Using tkFont
            for item in self.tree.get_children():
                cell_value = self.tree.set(item, col)
                cell_width = tkFont.Font().measure(cell_value)
                if cell_width > max_width:
                    max_width = cell_width
            self.tree.column(col, width=max_width + 10)  # Add a little padding

    def download_to_xlsx(self, event=None):
        # Create a new workbook and select the active worksheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reorder Quantities"

        # Write the headers to the worksheet
        headers = ['Parameter'] + self.month_names
        ws.append(headers)

        # Write the data from the TreeView to the worksheet
        for row in self.tree.get_children():
            row_data = self.tree.item(row)['values']
            ws.append(row_data)

        # Save the workbook to the specified path
        file_path = r"C:\Users\Frank\Desktop\Reorder_Quantities1.xlsx"
        wb.save(file_path)
        messagebox.showinfo("Download Complete", f"File has been saved to {file_path}")

    def paste_from_clipboard(self, row):
        try:
            clipboard_data = self.clipboard_get()
            values = clipboard_data.strip().split()  # Split by whitespace

            # Debugging: print the clipboard data and values
            print("Clipboard data:", clipboard_data)
            print("Parsed values:", values)

            num_months = int(self.months_ahead.get())

            expected_values = num_months  # Only values for the selected row
            if len(values) != expected_values:
                messagebox.showerror("Error", f"Clipboard data does not match the expected number of input fields. Expected {expected_values} values, got {len(values)}.")
                return

            for col in range(num_months):
                value = values[col]
                print(f"Inserting {value} into entry {row}, {col}")  # Debugging print
                self.entries[row][col].delete(0, tk.END)
                self.entries[row][col].insert(0, value)
                print(f"Entry {row}, {col} now contains {self.entries[row][col].get()}")  # Additional debugging print
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while pasting data: {e}")

    def focus_next_widget(self, event):
        event.widget.tk_focusNext().focus()
        return "break"

if __name__ == "__main__":
    app = ReorderCalculator()
    app.mainloop()
