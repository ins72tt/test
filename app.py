import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import webbrowser
import re

def clean_column_name(name):
    return name.replace(' ', '_').replace('-', '_')

def extract_urls(text):
    return re.findall(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', text)

def load_excel_data():
    # Ask the user to select an Excel file
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])

    # Check if a file was selected
    if file_path:
        # Read data from the Excel file
        df = pd.read_excel(file_path)

        # Set pandas option to display seconds in datetime differences
        pd.set_option('display.float_format', lambda x: str(x))

        # Convert time values to datetime, including seconds if present
        df['Begintijd'] = pd.to_datetime(df['Begintijd'], format='%m/%d/%y %H:%M:%S', errors='coerce')
        df['Tijd van voltooien'] = pd.to_datetime(df['Tijd van voltooien'], format='%m/%d/%y %H:%M:%S', errors='coerce')

        # Calculate time spent for each row
        df['Time_Spent'] = df['Tijd van voltooien'] - df['Begintijd']

        # Translate column names to Dutch
        df = df.rename(columns={
            'Begintijd': 'Starttijd',
            'Tijd van voltooien': 'Voltooid op',
            'Time_Spent': 'Tijd besteed'
        })

        # Identify columns for treeview
        treeview_columns = ['ID', 'Starttijd', 'Voltooid op', 'Tijd besteed', 'E-mail', 'Naam']

        # Create the main window
        root = tk.Tk()
        root.title('Excel Data Viewer')

        # Configure row and column weights for resizing
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.columnconfigure(2, weight=1)
        root.rowconfigure(0, weight=1)

        # Create treeview
        tree = ttk.Treeview(root)
        tree['columns'] = [clean_column_name(col) for col in treeview_columns]
        tree.column("#0", width=0, stretch=tk.NO)
        tree.heading("#0", text="", anchor=tk.CENTER)

        # Add columns to treeview
        for col, clean_col in zip(treeview_columns, tree['columns']):
            tree.column(clean_col, anchor=tk.CENTER, width=120)
            tree.heading(clean_col, text=col, anchor=tk.CENTER)

        # Add data rows
        for index, row in df.iterrows():
            tree.insert("", tk.END, text=index, values=[row[col] for col in treeview_columns])

        tree.grid(row=0, column=0, sticky='nsew')

        # Image frame
        image_frame = tk.Frame(root, bg="white")  # Set background color to white
        image_frame.grid(row=0, column=2, sticky='nsew')

        # Dummy label for the title
        tk.Label(image_frame, text="Afbeeldingen", pady=1.5, bg="white").grid(row=0, column=1)

        # Link callback for URLs in the entire row
        def click_link(url):
            webbrowser.open_new(url)

        # Start the current_row from 1
        current_row = 1

        for index, row in df.iterrows():
            # Check if there are URLs in the row
            urls = [url.strip() for cell in row for url in extract_urls(str(cell)) if 'http' in url.lower()]

            if urls:
                pady_value = 0.30
                padx_value = 0  # Adjust the padding on the x-axis
                # Create labels only if there are URLs
                for i, url in enumerate(urls, start=1):
                    label = tk.Label(image_frame, text=f"Link {i}", fg="blue", cursor="hand2", bg="white")
                    label.grid(row=current_row, column=i - 1, pady=pady_value, padx=padx_value)
                    label.bind("<Button-1>", lambda e, url=url: click_link(url))

                # Increment the current_row for the next iteration
                current_row += 1
            else:
                # If no URLs in the row, add a dummy label with an empty space
                tk.Label(image_frame, text="", pady=0.30, bg="white").grid(row=current_row, column=0)
                current_row += 1

        root.mainloop()

# Create root window
root = tk.Tk()
root.title('Excel Data Viewer')

# Load Button
load_button = tk.Button(root, text="Load Excel File", command=load_excel_data)
load_button.grid(row=0, column=0, sticky='w', padx=10, pady=10)

# Run the Tkinter event loop
root.mainloop()
