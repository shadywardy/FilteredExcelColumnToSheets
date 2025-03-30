import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
from collections import OrderedDict  
import pyexcel as p



def set_window_icon(root):
    try:
        # For development (running as .py)
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
        if os.path.exists(icon_path):
            img = tk.PhotoImage(file=icon_path)
            root.iconphoto(False, img)
        else:
            # For compiled version (running as .exe)
            temp_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(temp_dir, "icon.png")
            img = tk.PhotoImage(file=icon_path)
            root.iconphoto(False, img)
    except Exception as e:
        print(f"Could not load icon: {e}")



class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor by Shady Wardy")
        set_window_icon(root)  
        
        self.root.geometry("800x700")  # Width x Height + Xpos + Ypos
        # self.root.resizable(False, True)  # Allow horizontal resize, prevent vertical resize
        
        
        
        self.root.wm_attributes("-topmost", True)  # Bring to front
        self.root.after_idle(lambda: self.root.wm_attributes("-topmost", False))



        # Variables
        self.file_path = tk.StringVar()
        self.column_var = tk.StringVar()
        self.file_name_var = tk.StringVar(value="output")
        self.file_type_var = tk.IntVar(value=2)  # Default to xlsx
        self.progress_var = tk.DoubleVar()  # Main processing progress
        self.import_progress_var = tk.DoubleVar()  # File import progress
        self.log_text = ""
        
        # Create GUI components
        self.create_widgets()
        
        
        
        
    def create_widgets(self):
        # Configure style for modern look
        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabelFrame', background='#f0f0f0', font=('Helvetica', 10, 'bold'))
        style.configure('TLabel', background='#f0f0f0', font=('Helvetica', 9))
        style.configure('TButton', font=('Helvetica', 9), padding=5)
        style.configure('TCombobox', padding=5)
        style.configure('TRadiobutton', background='#f0f0f0')
        style.map('TButton', background=[('active', '#e0e0e0')])

        # Main container with modern background
        main_container = ttk.Frame(self.root, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Create notebook for tabbed interface
        notebook = ttk.Notebook(main_container)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Main tab
        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="Main Settings")

        # File Selection with improved layout
        file_frame = ttk.LabelFrame(main_tab, text="1. Select Excel File")
        file_frame.pack(fill=tk.X, pady=5, padx=5, ipady=5, ipadx=5)

        file_entry = ttk.Entry(file_frame, textvariable=self.file_path, width=50, font=('Helvetica', 9))
        file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(file_frame, text="File Path:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_file, style='Accent.TButton')
        browse_btn.grid(row=0, column=2, padx=5, pady=5)

        # Progress bars with better visual
        progress_frame = ttk.LabelFrame(main_tab, text="Progress Indicators")
        progress_frame.pack(fill=tk.X, pady=5, padx=5, ipady=5, ipadx=5)

        # Import Progress
        ttk.Label(progress_frame, text="Import Progress:").pack(anchor='w', padx=5, pady=(5,0))
        self.import_progress_bar = ttk.Progressbar(progress_frame, variable=self.import_progress_var, maximum=100, style='green.Horizontal.TProgressbar')
        self.import_progress_bar.pack(fill=tk.X, padx=5, pady=(0,5))
        self.import_progress_label = ttk.Label(progress_frame, text="Ready to import", style='Status.TLabel')
        self.import_progress_label.pack(anchor='e', padx=5, pady=(0,5))

        # Processing Progress
        ttk.Label(progress_frame, text="Processing Progress:").pack(anchor='w', padx=5, pady=(5,0))
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, style='blue.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, padx=5, pady=(0,5))
        self.progress_label = ttk.Label(progress_frame, text="Ready", style='Status.TLabel')
        self.progress_label.pack(anchor='e', padx=5, pady=(0,5))

        # Settings frame for other options
        settings_frame = ttk.Frame(main_tab)
        settings_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Left settings column
        left_col = ttk.Frame(settings_frame)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # File Type Selection with better grouping
        type_frame = ttk.LabelFrame(left_col, text="2. Output File Type")
        type_frame.pack(fill=tk.X, pady=5, ipady=5, ipadx=5)

        type_btn_frame = ttk.Frame(type_frame)
        type_btn_frame.pack(pady=5)
        ttk.Radiobutton(type_btn_frame, text="Excel (.xlsx)", variable=self.file_type_var, value=2).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(type_btn_frame, text="Excel 97-2003 (.xls)", variable=self.file_type_var, value=1).pack(side=tk.LEFT, padx=10)

        # Column Selection with search capability
        column_frame = ttk.LabelFrame(left_col, text="3. Column to Filter")
        column_frame.pack(fill=tk.X, pady=5, ipady=5, ipadx=5)

        ttk.Label(column_frame, text="Available Columns:").pack(anchor='w', padx=5, pady=(5,0))
        self.column_combobox = ttk.Combobox(column_frame, textvariable=self.column_var, state="readonly", height=15)
        self.column_combobox.pack(fill=tk.X, padx=5, pady=(0,5), ipady=3)

        # Right settings column
        right_col = ttk.Frame(settings_frame)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        # Output File Name with placeholder
        name_frame = ttk.LabelFrame(right_col, text="4. Output File Name")
        name_frame.pack(fill=tk.X, pady=5, ipady=5, ipadx=5)

        ttk.Label(name_frame, text="Base Name (without extension):").pack(anchor='w', padx=5, pady=(5,0))
        name_entry = ttk.Entry(name_frame, textvariable=self.file_name_var)
        name_entry.pack(fill=tk.X, padx=5, pady=(0,5), ipady=3)

        # Save Location with visual feedback
        save_frame = ttk.LabelFrame(right_col, text="5. Save Location")
        save_frame.pack(fill=tk.X, pady=5, ipady=5, ipadx=5)

        save_btn = ttk.Button(save_frame, text="Choose Save Folder", command=self.choose_save_folder)
        save_btn.pack(pady=5)
        self.save_location_label = ttk.Label(save_frame, text="No folder selected", foreground='#666666')
        self.save_location_label.pack(pady=(0,5))

        # Log tab
        log_tab = ttk.Frame(notebook)
        notebook.add(log_tab, text="Log Output")

        # Enhanced log area with clear formatting
        log_frame = ttk.LabelFrame(log_tab, text="Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5, ipady=5, ipadx=5)

        self.log_area = ScrolledText(log_frame, height=15, wrap=tk.WORD, font=('Consolas', 9), 
                                   padx=10, pady=10, bg='#ffffff', fg='#333333')
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Button bar at bottom with improved styling
        button_bar = ttk.Frame(self.root, style='ButtonBar.TFrame')
        button_bar.pack(fill=tk.X, pady=(0,5), padx=10)

        # Create custom style for action buttons
        style.configure('Action.TButton', font=('Helvetica', 9, 'bold'), padding=8)

        self.process_button = ttk.Button(button_bar, text="PROCESS FILES", command=self.start_processing, 
                                       style='Action.TButton')
        self.process_button.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        ttk.Button(button_bar, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(button_bar, text="Try Another File", command=self.browse_file).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        exit_btn = ttk.Button(button_bar, text="Exit", command=self.root.quit, style='Exit.TButton')
        exit_btn.pack(side=tk.RIGHT, padx=5, fill=tk.X, expand=True)

        # Configure tags for colored log messages
        self.log_area.tag_config('success', foreground='#008000')
        self.log_area.tag_config('error', foreground='#ff0000')
        self.log_area.tag_config('warning', foreground='#ff8c00')
        self.log_area.tag_config('info', foreground='#000080')

        # Set focus to first field
        file_entry.focus_set()
            
        
    def reset_for_new_file(self):
        """Reset the interface to process a new file"""
        self.file_path.set("")
        self.column_var.set("")
        self.column_combobox['values'] = []
        self.progress_var.set(0)
        self.progress_label.config(text="Ready")
        self.import_progress_var.set(0)
        self.import_progress_label.config(text="Ready to import")
        self.log_message("\nReady to process a new file...")
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.load_columns()
            
    def choose_save_folder(self):
        folder_path = filedialog.askdirectory(title="Select Save Location")
        if folder_path:
            self.save_folder = folder_path
            self.save_location_label.config(text=folder_path)
            
    def load_columns(self):
        try:
            # Update import progress
            self.import_progress_var.set(10)
            self.import_progress_label.config(text="Loading file...")
            self.root.update_idletasks()
            
            df = pd.read_excel(self.file_path.get())
            
            # Update import progress
            self.import_progress_var.set(50)
            self.root.update_idletasks()
            
            columns = df.columns.tolist()
            self.column_combobox['values'] = columns
            if columns:
                self.column_var.set(columns[0])
            
            # Complete import progress
            self.import_progress_var.set(100)
            self.import_progress_label.config(text="File loaded successfully")
            self.log_message(f"Loaded file: {self.file_path.get()}\nFound columns: {', '.join(columns)}")
            
        except Exception as e:
            self.import_progress_var.set(0)
            self.import_progress_label.config(text="Import failed")
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")
            self.log_message(f"Error loading file: {str(e)}")
            
    def log_message(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        self.log_area.delete(1.0, tk.END)
        
    def start_processing(self):
        if not hasattr(self, 'save_folder'):
            messagebox.showwarning("Warning", "Please select a save location first")
            return
            
        if not self.file_path.get():
            messagebox.showwarning("Warning", "Please select an Excel file first")
            return
            
        if not self.column_var.get():
            messagebox.showwarning("Warning", "Please select a column to filter")
            return
            
        # Disable buttons during processing
        self.process_button.config(state=tk.DISABLED)
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Button) and widget != self.process_button:
                widget.config(state=tk.DISABLED)
                
        # Start processing in a separate thread
        processing_thread = threading.Thread(target=self.process_excel_file, daemon=True)
        processing_thread.start()
        
    def process_excel_file(self):
        try:
            input_file = self.file_path.get()
            filter_column = self.column_var.get()
            file_name = self.file_name_var.get()
            file_extension = '.xls' if self.file_type_var.get() == 1 else '.xlsx'
            
            self.log_message(f"\nStarting processing...")
            self.log_message(f"Input file: {input_file}")
            self.log_message(f"Filter column: {filter_column}")
            self.log_message(f"Output format: {file_extension}")
            
            df = pd.read_excel(input_file)
            original_columns = df.columns.tolist()  # Store original column order
            
            if filter_column not in df.columns:
                messagebox.showerror("Error", f"The column '{filter_column}' does not exist in the Excel file.")
                self.log_message(f"Error: Column '{filter_column}' not found")
                return
                
            unique_values = df[filter_column].unique()
            total_values = len(unique_values)
            self.log_message(f"Found {total_values} unique values to process")
            
            self.progress_var.set(0)
            self.progress_label.config(text="0%")
            
            for i, value in enumerate(unique_values, 1):
                filtered_df = df[df[filter_column] == value]
                clean_value = str(value).replace("/", "_").replace("\\", "_")
                output_file = os.path.join(self.save_folder, f"{file_name}_{clean_value}{file_extension}")

                # Reorder columns to match original sequence
                filtered_df = filtered_df[original_columns]
                
                if file_extension == '.xls':
                    # Create ordered records to preserve column sequence
                    records = []
                    for _, row in filtered_df.iterrows():
                        ordered_record = OrderedDict()
                        for col in original_columns:
                            ordered_record[col] = row[col]
                        records.append(ordered_record)
                    p.save_as(records=records, dest_file_name=output_file)
                else:
                    filtered_df.to_excel(output_file, index=False)
                    
                progress = (i / total_values) * 100
                self.progress_var.set(progress)
                self.progress_label.config(text=f"{int(progress)}%")
                self.log_message(f"Generated: {output_file}")
                self.root.update_idletasks()
                
            self.log_message("\nProcessing completed successfully!")
            messagebox.showinfo("Success", "All files have been generated successfully!")
            
        except Exception as e:
            self.log_message(f"\nError during processing: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
        finally:
            # Re-enable buttons
            for widget in self.root.winfo_children():
                if isinstance(widget, ttk.Button):
                    widget.config(state=tk.NORMAL)
                    
            self.progress_var.set(0)
            self.progress_label.config(text="Ready")

def main():
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()