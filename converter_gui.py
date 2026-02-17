#!/usr/bin/env python3
"""
PPTX to PDF Converter - GUI Application
Simple and intuitive interface for bulk PowerPoint to PDF conversion
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
from pathlib import Path
import sys
from convert_pptx_to_pdf import PPTXtoPDFConverter


class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPTX to PDF Converter")
        self.root.geometry("700x600")
        self.root.resizable(True, True)

        # Initialize converter
        self.converter = PPTXtoPDFConverter()

        # Queue for thread-safe GUI updates
        self.message_queue = queue.Queue()

        # Track conversion state
        self.is_converting = False

        # Setup UI
        self.setup_ui()

        # Check LibreOffice on startup
        self.check_libreoffice()

        # Start queue processor
        self.process_queue()

    def setup_ui(self):
        """Create the user interface"""

        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="PPTX to PDF Converter",
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=0, pady=(0, 10))

        subtitle_label = ttk.Label(
            main_frame,
            text="Convert PowerPoint presentations to PDF with high quality",
            font=('Arial', 9)
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))

        # Option buttons frame
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="15")
        options_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        options_frame.columnconfigure(0, weight=1)

        # Option 1: Convert Single File
        btn_single = ttk.Button(
            options_frame,
            text="📄 Convert Single File",
            command=self.convert_single_file,
            width=40
        )
        btn_single.grid(row=0, column=0, pady=5, padx=10, sticky=(tk.W, tk.E))

        label_single = ttk.Label(
            options_frame,
            text="Select one PPTX file to convert to PDF",
            font=('Arial', 8),
            foreground='gray'
        )
        label_single.grid(row=1, column=0, pady=(0, 10), padx=10)

        # Option 2: Convert Multiple Files
        btn_multiple = ttk.Button(
            options_frame,
            text="📑 Convert Multiple Files",
            command=self.convert_multiple_files,
            width=40
        )
        btn_multiple.grid(row=2, column=0, pady=5, padx=10, sticky=(tk.W, tk.E))

        label_multiple = ttk.Label(
            options_frame,
            text="Select multiple PPTX files to convert in batch",
            font=('Arial', 8),
            foreground='gray'
        )
        label_multiple.grid(row=3, column=0, pady=(0, 10), padx=10)

        # Option 3: Convert Entire Folder
        btn_folder = ttk.Button(
            options_frame,
            text="📁 Convert Entire Folder",
            command=self.convert_folder,
            width=40
        )
        btn_folder.grid(row=4, column=0, pady=5, padx=10, sticky=(tk.W, tk.E))

        label_folder = ttk.Label(
            options_frame,
            text="Convert all PPTX files in a selected folder",
            font=('Arial', 8),
            foreground='gray'
        )
        label_folder.grid(row=5, column=0, pady=(0, 10), padx=10)

        # Output settings frame
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)

        # Output directory selector
        ttk.Label(output_frame, text="Output Directory:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))

        self.output_dir_var = tk.StringVar(value="(Same as input)")
        output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, state='readonly')
        output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        btn_browse_output = ttk.Button(
            output_frame,
            text="Browse...",
            command=self.select_output_directory,
            width=12
        )
        btn_browse_output.grid(row=0, column=2)

        btn_reset_output = ttk.Button(
            output_frame,
            text="Reset",
            command=self.reset_output_directory,
            width=10
        )
        btn_reset_output.grid(row=0, column=3, padx=(5, 0))

        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Conversion Progress", padding="10")
        progress_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Status text area
        self.status_text = scrolledtext.ScrolledText(
            progress_frame,
            height=12,
            wrap=tk.WORD,
            font=('Consolas', 9)
        )
        self.status_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        progress_frame.rowconfigure(1, weight=1)

        # Bottom frame with action buttons
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        bottom_frame.columnconfigure(0, weight=1)

        # Clear log button
        btn_clear = ttk.Button(
            bottom_frame,
            text="Clear Log",
            command=self.clear_log
        )
        btn_clear.grid(row=0, column=0, sticky=tk.W)

        # LibreOffice status
        self.libreoffice_status = ttk.Label(
            bottom_frame,
            text="Checking LibreOffice...",
            font=('Arial', 8)
        )
        self.libreoffice_status.grid(row=0, column=1, padx=20)

    def check_libreoffice(self):
        """Check if LibreOffice is installed"""
        if self.converter.libreoffice_path:
            self.libreoffice_status.config(
                text=f"✓ LibreOffice Ready",
                foreground='green'
            )
            self.log_message(f"LibreOffice found: {self.converter.libreoffice_path}\n")
        else:
            self.libreoffice_status.config(
                text="✗ LibreOffice Not Found",
                foreground='red'
            )
            self.log_message("ERROR: LibreOffice not installed!\n", 'error')
            self.log_message("Please install LibreOffice from: https://www.libreoffice.org/download/download/\n", 'error')

    def log_message(self, message, tag='normal'):
        """Add message to status text area"""
        self.status_text.insert(tk.END, message)
        if tag == 'error':
            # Get the last line and apply red color
            self.status_text.tag_add('error', 'end-2l', 'end-1l')
            self.status_text.tag_config('error', foreground='red')
        elif tag == 'success':
            self.status_text.tag_add('success', 'end-2l', 'end-1l')
            self.status_text.tag_config('success', foreground='green')
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """Clear the status text area"""
        self.status_text.delete(1.0, tk.END)
        self.progress_var.set(0)

    def select_output_directory(self):
        """Select output directory"""
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_var.set(directory)
            self.log_message(f"Output directory set to: {directory}\n")

    def reset_output_directory(self):
        """Reset output directory to default"""
        self.output_dir_var.set("(Same as input)")
        self.log_message("Output directory reset to input location\n")

    def get_output_directory(self):
        """Get the output directory or None if using default"""
        output = self.output_dir_var.get()
        return None if output == "(Same as input)" else output

    def convert_single_file(self):
        """Option 1: Convert a single PPTX file"""
        if self.is_converting:
            messagebox.showwarning("Busy", "A conversion is already in progress!")
            return

        if not self.converter.libreoffice_path:
            messagebox.showerror("Error", "LibreOffice is not installed!\n\nPlease install LibreOffice first:\nhttps://www.libreoffice.org/download/download/")
            return

        file_path = filedialog.askopenfilename(
            title="Select PPTX File",
            filetypes=[
                ("PowerPoint Files", "*.pptx *.ppt"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            self.clear_log()
            self.log_message(f"Selected: {file_path}\n")
            self.log_message("=" * 60 + "\n")

            # Run conversion in separate thread
            thread = threading.Thread(
                target=self.run_conversion,
                args=([file_path],),
                daemon=True
            )
            thread.start()

    def convert_multiple_files(self):
        """Option 2: Convert multiple PPTX files"""
        if self.is_converting:
            messagebox.showwarning("Busy", "A conversion is already in progress!")
            return

        if not self.converter.libreoffice_path:
            messagebox.showerror("Error", "LibreOffice is not installed!\n\nPlease install LibreOffice first:\nhttps://www.libreoffice.org/download/download/")
            return

        file_paths = filedialog.askopenfilenames(
            title="Select PPTX Files",
            filetypes=[
                ("PowerPoint Files", "*.pptx *.ppt"),
                ("All Files", "*.*")
            ]
        )

        if file_paths:
            self.clear_log()
            self.log_message(f"Selected {len(file_paths)} file(s)\n")
            self.log_message("=" * 60 + "\n")

            # Run conversion in separate thread
            thread = threading.Thread(
                target=self.run_conversion,
                args=(list(file_paths),),
                daemon=True
            )
            thread.start()

    def convert_folder(self):
        """Option 3: Convert all PPTX files in a folder"""
        if self.is_converting:
            messagebox.showwarning("Busy", "A conversion is already in progress!")
            return

        if not self.converter.libreoffice_path:
            messagebox.showerror("Error", "LibreOffice is not installed!\n\nPlease install LibreOffice first:\nhttps://www.libreoffice.org/download/download/")
            return

        folder_path = filedialog.askdirectory(title="Select Folder Containing PPTX Files")

        if folder_path:
            # Count PPTX files in folder
            folder = Path(folder_path)
            pptx_files = list(folder.rglob('*.pptx')) + list(folder.rglob('*.ppt'))
            pptx_files += list(folder.rglob('*.PPTX')) + list(folder.rglob('*.PPT'))

            if not pptx_files:
                messagebox.showwarning("No Files", f"No PPTX files found in:\n{folder_path}")
                return

            self.clear_log()
            self.log_message(f"Selected folder: {folder_path}\n")
            self.log_message(f"Found {len(pptx_files)} PPTX file(s)\n")
            self.log_message("=" * 60 + "\n")

            # Run conversion in separate thread
            thread = threading.Thread(
                target=self.run_conversion,
                args=([folder_path],),
                daemon=True
            )
            thread.start()

    def run_conversion(self, input_paths):
        """Run the conversion in a separate thread"""
        self.is_converting = True
        output_dir = self.get_output_directory()

        try:
            # Collect all files
            files_to_convert = []
            for path in input_paths:
                p = Path(path)
                if p.is_file() and p.suffix.lower() in ['.pptx', '.ppt']:
                    files_to_convert.append(p)
                elif p.is_dir():
                    files_to_convert.extend(p.rglob('*.pptx'))
                    files_to_convert.extend(p.rglob('*.ppt'))
                    files_to_convert.extend(p.rglob('*.PPTX'))
                    files_to_convert.extend(p.rglob('*.PPT'))

            if not files_to_convert:
                self.message_queue.put(('log', "No PPTX files to convert.\n", 'error'))
                return

            total_files = len(files_to_convert)
            success_count = 0
            failed_count = 0

            self.message_queue.put(('log', f"\nStarting conversion of {total_files} file(s)...\n\n", 'normal'))

            for i, file_path in enumerate(files_to_convert, 1):
                # Update progress
                progress = (i - 1) / total_files * 100
                self.message_queue.put(('progress', progress))

                file_size = file_path.stat().st_size / (1024 * 1024)  # MB
                self.message_queue.put(('log', f"[{i}/{total_files}] Converting: {file_path.name} ({file_size:.2f} MB)\n", 'normal'))

                # Convert file
                try:
                    result = self.converter.convert_file(str(file_path), output_dir, verbose=False)

                    if result:
                        success_count += 1
                        self.message_queue.put(('log', f"✓ Success: {file_path.name}\n", 'success'))
                    else:
                        failed_count += 1
                        self.message_queue.put(('log', f"✗ Failed: {file_path.name}\n", 'error'))
                except Exception as e:
                    failed_count += 1
                    self.message_queue.put(('log', f"✗ Error: {file_path.name} - {str(e)}\n", 'error'))

                # Update progress
                progress = i / total_files * 100
                self.message_queue.put(('progress', progress))

            # Final summary
            self.message_queue.put(('log', "\n" + "=" * 60 + "\n", 'normal'))
            self.message_queue.put(('log', "Conversion Summary:\n", 'normal'))
            self.message_queue.put(('log', f"  Total files: {total_files}\n", 'normal'))
            self.message_queue.put(('log', f"  Successful: {success_count}\n", 'success'))
            if failed_count > 0:
                self.message_queue.put(('log', f"  Failed: {failed_count}\n", 'error'))
            self.message_queue.put(('log', "\nConversion complete!\n", 'success'))

            # Show completion message
            if failed_count == 0:
                self.message_queue.put(('msgbox', ('success', 'Conversion Complete', f'Successfully converted {success_count} file(s)!')))
            else:
                self.message_queue.put(('msgbox', ('warning', 'Conversion Complete', f'Converted {success_count} file(s)\n{failed_count} file(s) failed')))

        except Exception as e:
            self.message_queue.put(('log', f"\nError during conversion: {str(e)}\n", 'error'))
            self.message_queue.put(('msgbox', ('error', 'Conversion Error', str(e))))

        finally:
            self.is_converting = False

    def process_queue(self):
        """Process messages from the conversion thread"""
        try:
            while True:
                msg_type, *args = self.message_queue.get_nowait()

                if msg_type == 'log':
                    self.log_message(args[0], args[1] if len(args) > 1 else 'normal')
                elif msg_type == 'progress':
                    self.progress_var.set(args[0])
                elif msg_type == 'msgbox':
                    box_type, title, message = args[0]
                    if box_type == 'success':
                        messagebox.showinfo(title, message)
                    elif box_type == 'warning':
                        messagebox.showwarning(title, message)
                    elif box_type == 'error':
                        messagebox.showerror(title, message)

        except queue.Empty:
            pass

        # Schedule next check
        self.root.after(100, self.process_queue)


def main():
    """Launch the GUI application"""
    root = tk.Tk()
    app = ConverterGUI(root)

    # Set window icon (optional - would need an icon file)
    try:
        root.iconbitmap('icon.ico')
    except:
        pass

    # Center window on screen
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == '__main__':
    main()
