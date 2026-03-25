import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, Enum
from sqlalchemy.orm import declarative_base, sessionmaker
import enum
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import csv
import os

# Define modern color scheme with better cross-platform readability
COLORS = {
    'primary': '#2c3e50',      # Dark blue-gray
    'secondary': '#2980b9',    # Darker blue for better contrast
    'accent': '#c0392b',       # Darker red for better contrast
    'success': '#27ae60',      # Darker green for better contrast
    'warning': '#f39c12',      # Darker yellow for better contrast
    'background': '#f5f5f5',   # Lighter gray for better contrast with text
    'text': '#000000',         # Black for maximum readability
    'text_light': '#2c3e50',   # Darker blue-gray for better contrast
    'white': '#ffffff',        # White
    'border': '#95a5a6',        # Darker medium gray for better visibility
    'dark': '#7f8c8d',          # Darker gray
    'table_header': '#d5d5d5',  # Light gray for table headers
    'table_row_alt': '#eaeaea'  # Alternate row color for better readability
}

# Configure ttk styles
def configure_styles():
    style = ttk.Style()
    # Native Windows themes often ignore custom fg/bg maps (white-on-white hover). Clam honors them.
    if "clam" in style.theme_names():
        style.theme_use("clam")

    # Configure main window style
    style.configure("Main.TFrame", background=COLORS["background"])

    # Notebook: unselected = gray tab + dark text; selected = white panel + dark text (readable on all platforms)
    style.configure("TNotebook", background=COLORS["background"], borderwidth=0)
    style.configure(
        "TNotebook.Tab",
        padding=[12, 6],
        background=COLORS["table_header"],
        foreground=COLORS["text"],
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", COLORS["white"])],
        foreground=[("selected", COLORS["text"])],
    )

    # Primary actions: keep label light on dark for every interactive state (hover/press/focus)
    style.configure(
        "Primary.TButton",
        padding=10,
        background=COLORS["primary"],
        foreground=COLORS["white"],
        borderwidth=1,
        focuscolor=COLORS["primary"],
    )
    style.map(
        "Primary.TButton",
        background=[
            ("disabled", COLORS["border"]),
            ("pressed", COLORS["secondary"]),
            ("active", COLORS["secondary"]),
            ("!disabled", COLORS["primary"]),
        ],
        foreground=[
            ("disabled", COLORS["text"]),
            ("pressed", COLORS["white"]),
            ("active", COLORS["white"]),
            ("!disabled", COLORS["white"]),
        ],
    )

    # Default buttons: avoid light-on-light when the theme highlights
    style.configure("TButton", background=COLORS["background"])
    style.map(
        "TButton",
        background=[
            ("pressed", COLORS["table_header"]),
            ("active", COLORS["table_header"]),
            ("!disabled", COLORS["background"]),
        ],
        foreground=[
            ("disabled", COLORS["dark"]),
            ("!disabled", COLORS["text"]),
        ],
    )
    
    # Configure entry styles
    style.configure('TEntry',
        padding=5,
        fieldbackground=COLORS['white'],
        borderwidth=1
    )
    
    # Configure label styles
    style.configure('TLabel',
        background=COLORS['background'],
        foreground=COLORS['text'],
        padding=5
    )
    
    # Configure treeview styles with improved contrast
    style.configure('Treeview',
        background=COLORS['white'],
        foreground=COLORS['text'],
        fieldbackground=COLORS['white'],
        rowheight=25
    )
    
    # Configure treeview headings with more visible styling
    style.configure('Treeview.Heading',
        background=COLORS['table_header'],  # Light gray background for better text contrast
        foreground=COLORS['text'],  # Black text for maximum readability
        padding=5,
        relief='raised',  # Add relief to make headings stand out
        borderwidth=1,
        font=('Helvetica', 10, 'bold')  # Make headings bold for better visibility
    )
    
    # Ensure the heading style is properly applied by explicitly setting it
    style.layout('Treeview.Heading', [
        ('Treeview.Heading.cell', {
            'sticky': 'nswe',
            'children': [
                ('Treeview.Heading.border', {
                    'sticky': 'nswe',
                    'border': '1',
                    'children': [
                        ('Treeview.Heading.padding', {
                            'sticky': 'nswe',
                            'children': [
                                ('Treeview.Heading.image', {'side': 'right', 'sticky': ''}),
                                ('Treeview.Heading.text', {'sticky': 'we'})
                            ]
                        })
                    ]
                })
            ]
        })
    ])
    
    # Add hover effect for headings
    style.map('Treeview.Heading',
        background=[('active', '#b0b0b0')],  # Slightly darker gray when active
        foreground=[('active', '#000000')]
    )
    style.map('Treeview',
        background=[('selected', COLORS['secondary'])],
        foreground=[('selected', COLORS['white'])]
    )
    
    # Add alternating row colors for better readability - handled in the treeview creation
    # as ttk.Treeview doesn't directly support alternating row colors through style mapping
    
    # Configure combobox styles
    style.configure('TCombobox',
        background=COLORS['white'],
        fieldbackground=COLORS['white'],
        selectbackground=COLORS['secondary'],
        selectforeground=COLORS['white']
    )
    
    # Configure frame styles
    style.configure('Card.TFrame',
        background=COLORS['white'],
        relief='solid',
        borderwidth=1
    )
    
    # Configure separator style
    style.configure('TSeparator',
        background=COLORS['border']
    )


def master_detail_scroll_setup(container, title_var, initial_title):
    """Title + scrollable form column (same pattern as Planned Changes detail pane)."""
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)
    title_var.set(initial_title)
    ttk.Label(
        container,
        textvariable=title_var,
        font=("Helvetica", 11, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))
    canvas = tk.Canvas(container, highlightthickness=0)
    vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=vsb.set)
    form = ttk.Frame(canvas, padding=(0, 0, 6, 0))
    inner_win = canvas.create_window((0, 0), window=form, anchor=tk.NW)

    def sync_scroll(_event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_cfg(event):
        canvas.itemconfigure(inner_win, width=event.width)

    form.bind("<Configure>", sync_scroll)
    canvas.bind("<Configure>", on_canvas_cfg)
    canvas.grid(row=1, column=0, sticky=tk.NSEW)
    vsb.grid(row=1, column=1, sticky=tk.NS)

    def on_mw(event):
        if canvas.winfo_containing(event.x_root, event.y_root):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

    canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", on_mw))
    canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
    return canvas, form


# Create database engine
engine = create_engine('sqlite:///forecast_tool.db')
Base = declarative_base()
Session = sessionmaker(bind=engine)

def get_session():
    return Session()

def verify_db_connection():
    """Verify database connection and reset it if necessary"""
    try:
        session = get_session()
        # Try a simple query
        session.query(Settings).first()
        session.close()
        return True
    except Exception as e:
        print(f"Database connection error: {str(e)}")
        try:
            # Try to reset the connection
            global engine, Session
            engine = create_engine('sqlite:///forecast_tool.db')
            Session = sessionmaker(bind=engine)
            # Create tables if they don't exist
            Base.metadata.create_all(engine)
            return True
        except Exception as e:
            print(f"Failed to reset database connection: {str(e)}")
            return False

# Enums
class EmploymentType(enum.Enum):
    FTE = "FTE"
    CONTRACTOR = "CONTRACTOR"

class ChangeType(enum.Enum):
    NEW_HIRE = "New Hire"
    CONVERSION = "Conversion"
    TERMINATION = "Termination"

# Database Models
class Employee(Base):
    __tablename__ = 'employees'
    
    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    manager_code = Column(String, nullable=False)
    cost_center = Column(String, nullable=False)
    employment_type = Column(String, nullable=False)
    work_code = Column(String, nullable=True)  # Added work code field
    start_date = Column(Date, nullable=False)
    end_date = Column(Date)

class ProjectAllocation(Base):
    __tablename__ = 'project_allocations'
    
    id = Column(Integer, primary_key=True)
    manager_code = Column(String, nullable=False)
    cost_center = Column(String, nullable=False)
    work_code = Column(String, nullable=False)
    year = Column(Integer, nullable=False)
    jan = Column(Float, default=0)
    feb = Column(Float, default=0)
    mar = Column(Float, default=0)
    apr = Column(Float, default=0)
    may = Column(Float, default=0)
    jun = Column(Float, default=0)
    jul = Column(Float, default=0)
    aug = Column(Float, default=0)
    sep = Column(Float, default=0)
    oct = Column(Float, default=0)
    nov = Column(Float, default=0)
    dec = Column(Float, default=0)

class Settings(Base):
    __tablename__ = 'settings'
    
    id = Column(Integer, primary_key=True)
    fte_hours = Column(Float, default=34.5)
    contractor_hours = Column(Float, default=39.0)

class GA01Week(Base):
    __tablename__ = 'ga01_weeks'
    
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    weeks = Column(Float, nullable=False)

class PlannedChange(Base):
    __tablename__ = 'planned_changes'
    
    id = Column(Integer, primary_key=True)
    description = Column(String, nullable=False)
    change_type = Column(String, nullable=False)
    effective_date = Column(Date, nullable=False)
    employee_id = Column(Integer)
    target_type = Column(String)
    name = Column(String)
    team = Column(String)
    manager_code = Column(String)
    cost_center = Column(String)
    employment_type = Column(String)
    status = Column(String, nullable=False)

class Forecast(Base):
    __tablename__ = 'forecasts'
    
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    cost_center = Column(String, nullable=False)
    manager_code = Column(String, nullable=False)
    work_code = Column(String, nullable=False)
    jan = Column(Float, default=0)
    feb = Column(Float, default=0)
    mar = Column(Float, default=0)
    apr = Column(Float, default=0)
    may = Column(Float, default=0)
    jun = Column(Float, default=0)
    jul = Column(Float, default=0)
    aug = Column(Float, default=0)
    sep = Column(Float, default=0)
    oct = Column(Float, default=0)
    nov = Column(Float, default=0)
    dec = Column(Float, default=0)
    total_hours = Column(Float, default=0)

# Create tables
Base.metadata.create_all(engine)

class EmployeeTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self._record_id = None
        self._detail_canvas = None
        self._suppress_tree_select = False
        self.create_widgets()

    def create_widgets(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Add Employee", command=self.add_employee).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Employee", command=self.edit_employee).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Employee", command=self.delete_employee).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Import Employees", command=self.import_employees).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_employees).pack(side=tk.LEFT, padx=2)

        self.paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=2)

        self.tree_frame = ttk.Frame(left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree_yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=("ID", "Name", "Manager Code", "Cost Center", "Type", "Start Date", "End Date"),
            show="headings",
            yscrollcommand=scrollbar.set,
        )
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.heading("ID", text="ID")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Manager Code", text="Manager Code")
        self.tree.heading("Cost Center", text="Cost Center")
        self.tree.heading("Type", text="Type")
        self.tree.heading("Start Date", text="Start Date")
        self.tree.heading("End Date", text="End Date")

        self.tree.column("ID", width=50)
        self.tree.column("Name", width=180)
        self.tree.column("Manager Code", width=90)
        self.tree.column("Cost Center", width=90)
        self.tree.column("Type", width=90)
        self.tree.column("Start Date", width=90)
        self.tree.column("End Date", width=90)

        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        right_outer = ttk.Frame(self.paned, padding=(8, 0, 0, 0))
        self.paned.add(right_outer, weight=1)

        self.detail_title_var = tk.StringVar(value="Employee")
        self._detail_canvas, form = master_detail_scroll_setup(right_outer, self.detail_title_var, "Employee")

        form.columnconfigure(1, weight=1)
        details = ttk.LabelFrame(form, text="Employee details", padding=8)
        details.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=4)
        details.columnconfigure(1, weight=1)

        ttk.Label(details, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.name_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.name_var, width=36).grid(row=0, column=1, sticky=tk.EW, pady=4)

        ttk.Label(details, text="Manager Code:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.manager_code_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.manager_code_var, width=22).grid(row=1, column=1, sticky=tk.W, pady=4)

        ttk.Label(details, text="Cost Center:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.cost_center_var, width=22).grid(row=2, column=1, sticky=tk.W, pady=4)

        ttk.Label(details, text="Work Code:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.work_code_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.work_code_var, width=22).grid(row=3, column=1, sticky=tk.W, pady=4)

        ttk.Label(details, text="Employment Type:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.employment_type_var = tk.StringVar()
        employment_types = [et.value for et in EmploymentType]
        ttk.Combobox(
            details, textvariable=self.employment_type_var, values=employment_types, width=20, state="readonly"
        ).grid(row=4, column=1, sticky=tk.W, pady=4)

        ttk.Label(details, text="Start Date (mm/dd/yy):").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.start_date_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.start_date_var, width=22).grid(row=5, column=1, sticky=tk.W, pady=4)

        ttk.Label(details, text="End Date (mm/dd/yy):").grid(row=6, column=0, sticky=tk.W, pady=4)
        self.end_date_var = tk.StringVar()
        ttk.Entry(details, textvariable=self.end_date_var, width=22).grid(row=6, column=1, sticky=tk.W, pady=4)

        button_bar = ttk.Frame(right_outer)
        button_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(10, 0))
        ttk.Button(button_bar, text="Save", command=self._save_detail, width=10).pack(side=tk.RIGHT, padx=4)
        ttk.Button(button_bar, text="Cancel", command=self._cancel_detail, width=10).pack(side=tk.RIGHT, padx=4)

        self.bind("<Destroy>", self._emp_on_destroy)
        self.after_idle(self._emp_set_initial_sash)
        self.load_employees()
        self._clear_detail_form(new_mode=True)

    def tree_yview(self, *args):
        self.tree.yview(*args)

    def _emp_set_initial_sash(self):
        try:
            self.paned.sashpos(0, 560)
        except tk.TclError:
            pass

    def _emp_on_destroy(self, event):
        if event.widget is not self:
            return
        try:
            if self._detail_canvas is not None:
                self._detail_canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass

    def _clear_detail_form(self, new_mode=False):
        self._record_id = None
        self.detail_title_var.set("New employee" if new_mode else "Employee")
        self.name_var.set("")
        self.manager_code_var.set("")
        self.cost_center_var.set("")
        self.work_code_var.set("")
        self.employment_type_var.set("")
        self.start_date_var.set(datetime.now().strftime("%m/%d/%y"))
        self.end_date_var.set("")

    def _apply_employee_to_form(self, emp):
        self._record_id = emp.id
        self.detail_title_var.set(f"Employee #{emp.id}")
        self.name_var.set(emp.name or "")
        self.manager_code_var.set(emp.manager_code or "")
        self.cost_center_var.set(emp.cost_center or "")
        self.work_code_var.set(emp.work_code or "")
        self.employment_type_var.set(emp.employment_type or "")
        self.start_date_var.set(emp.start_date.strftime("%m/%d/%y") if emp.start_date else "")
        self.end_date_var.set(emp.end_date.strftime("%m/%d/%y") if emp.end_date else "")

    def _load_employee_by_id(self, emp_id):
        session = get_session()
        try:
            employee = session.query(Employee).filter(Employee.id == emp_id).first()
            if employee:
                self._apply_employee_to_form(employee)
            else:
                messagebox.showerror("Error", "Employee not found.")
                self._clear_detail_form(new_mode=True)
        finally:
            session.close()

    def _on_tree_select(self, _event=None):
        if self._suppress_tree_select:
            return
        selected = self.tree.selection()
        if not selected:
            return
        emp_id = self.tree.item(selected[0])["values"][0]
        self._load_employee_by_id(emp_id)

    def _collect_detail_payload(self):
        if not self.name_var.get().strip():
            raise ValueError("Name is required")
        if not self.manager_code_var.get().strip():
            raise ValueError("Manager Code is required")
        if not self.cost_center_var.get().strip():
            raise ValueError("Cost Center is required")
        if not self.employment_type_var.get():
            raise ValueError("Employment Type is required")
        if not self.start_date_var.get().strip():
            raise ValueError("Start Date is required")
        start_date = datetime.strptime(self.start_date_var.get(), "%m/%d/%y").date()
        end_date = None
        if self.end_date_var.get().strip():
            end_date = datetime.strptime(self.end_date_var.get(), "%m/%d/%y").date()
            if end_date <= start_date:
                raise ValueError("End Date must be after Start Date")
        return {
            "name": self.name_var.get().strip(),
            "manager_code": self.manager_code_var.get().strip(),
            "cost_center": self.cost_center_var.get().strip(),
            "work_code": self.work_code_var.get().strip(),
            "employment_type": self.employment_type_var.get(),
            "start_date": start_date,
            "end_date": end_date,
        }

    def _save_detail(self):
        try:
            data = self._collect_detail_payload()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return

        try:
            session = get_session()
            if self._record_id is None:
                employee = Employee(
                    name=data["name"],
                    manager_code=data["manager_code"],
                    cost_center=data["cost_center"],
                    work_code=data["work_code"],
                    employment_type=data["employment_type"],
                    start_date=data["start_date"],
                    end_date=data["end_date"],
                )
                session.add(employee)
                session.commit()
                new_id = employee.id
                session.close()
                self.load_employees()
                self._select_tree_row_by_id(new_id)
                messagebox.showinfo("Success", "Employee added successfully.")
            else:
                employee = session.query(Employee).filter(Employee.id == self._record_id).first()
                if not employee:
                    session.close()
                    messagebox.showerror("Error", "Employee not found.")
                    return
                employee.name = data["name"]
                employee.manager_code = data["manager_code"]
                employee.cost_center = data["cost_center"]
                employee.work_code = data["work_code"]
                employee.employment_type = data["employment_type"]
                employee.start_date = data["start_date"]
                employee.end_date = data["end_date"]
                session.commit()
                rid = self._record_id
                session.close()
                self.load_employees()
                self._select_tree_row_by_id(rid)
                messagebox.showinfo("Success", "Employee updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save employee: {str(e)}")

    def _select_tree_row_by_id(self, emp_id):
        self._suppress_tree_select = True
        try:
            for item in self.tree.get_children():
                if self.tree.item(item)["values"][0] == emp_id:
                    self.tree.selection_set(item)
                    self.tree.see(item)
                    self._load_employee_by_id(emp_id)
                    return
        finally:
            self._suppress_tree_select = False

    def _cancel_detail(self):
        if self._record_id is not None:
            self._load_employee_by_id(self._record_id)
        else:
            sel = self.tree.selection()
            if sel:
                self._on_tree_select()
            else:
                self._clear_detail_form(new_mode=True)

    def load_employees(self):
        """Load employees from database"""
        keep_id = self._record_id
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            session = get_session()
            employees = session.query(Employee).all()
            count = 0
            for emp in employees:
                item_id = self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        emp.id,
                        emp.name,
                        emp.manager_code,
                        emp.cost_center,
                        emp.employment_type,
                        emp.start_date.strftime("%m/%d/%y") if emp.start_date else "",
                        emp.end_date.strftime("%m/%d/%y") if emp.end_date else "",
                    ),
                )
                self.tree.item(item_id, tags=("evenrow",) if count % 2 == 1 else ("oddrow",))
                count += 1
            self.tree.tag_configure("oddrow", background=COLORS["white"])
            self.tree.tag_configure("evenrow", background=COLORS["table_row_alt"])
            session.close()
            if keep_id is not None:
                if any(self.tree.item(i)["values"][0] == keep_id for i in self.tree.get_children()):
                    self._select_tree_row_by_id(keep_id)
                else:
                    self._clear_detail_form(new_mode=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employees: {str(e)}")

    def add_employee(self):
        self._suppress_tree_select = True
        try:
            self.tree.selection_remove(self.tree.selection())
        except tk.TclError:
            pass
        self._suppress_tree_select = False
        self._clear_detail_form(new_mode=True)

    def edit_employee(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select an employee to edit.")
            return
        emp_id = self.tree.item(selected[0])["values"][0]
        self._load_employee_by_id(emp_id)

    def delete_employee(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select an employee to delete.")
            return
        emp_id = self.tree.item(selection[0])["values"][0]
        if not messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete employee with ID {emp_id}?"):
            return
        try:
            session = get_session()
            employee = session.query(Employee).filter(Employee.id == emp_id).first()
            if employee:
                session.delete(employee)
                session.commit()
                if self._record_id == emp_id:
                    self._clear_detail_form(new_mode=True)
                self.load_employees()
                messagebox.showinfo("Success", f"Employee with ID {emp_id} deleted successfully.")
            else:
                messagebox.showerror("Error", "Employee not found.")
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete employee: {str(e)}")

    def import_employees(self):
        """Import employees from a CSV file"""
        # Ask user to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if not file_path:
            return  # User cancelled
        
        try:
            # Read CSV file with BOM handling
            with open(file_path, 'r', newline='', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                
                # Check required columns
                required_columns = ['name', 'manager_code', 'cost_center', 'employment_type', 'start_date']
                headers = reader.fieldnames
                
                if not all(col in headers for col in required_columns):
                    missing = [col for col in required_columns if col not in headers]
                    messagebox.showerror("Error", f"CSV file is missing required columns: {', '.join(missing)}\n\n"
                                       f"Required columns are: {', '.join(required_columns)}")
                    return
                
                # Begin import
                session = get_session()
                imported_count = 0
                error_count = 0
                error_messages = []
                
                for row in reader:
                    try:
                        # Parse dates
                        try:
                            start_date = datetime.strptime(row['start_date'], "%m/%d/%y").date()
                        except ValueError:
                            raise ValueError(f"Invalid start date format for {row['name']}: {row['start_date']}. Use MM/DD/YY.")
                        
                        end_date = None
                        if 'end_date' in row and row['end_date']:
                            try:
                                end_date = datetime.strptime(row['end_date'], "%m/%d/%y").date()
                            except ValueError:
                                raise ValueError(f"Invalid end date format for {row['name']}: {row['end_date']}. Use MM/DD/YY.")
                        
                        # Create employee
                        employee = Employee(
                            name=row['name'],
                            manager_code=row['manager_code'],
                            cost_center=row['cost_center'],
                            employment_type=row['employment_type'],
                            start_date=start_date,
                            end_date=end_date
                        )
                        
                        session.add(employee)
                        imported_count += 1
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"Row {reader.line_num}: {str(e)}")
                        if error_count >= 5:  # Limit error messages to avoid huge dialog
                            error_messages.append("(Additional errors not shown)")
                            break
                
                # Commit changes
                if imported_count > 0:
                    session.commit()
                
                session.close()
                
                # Show results
                if error_count > 0:
                    messagebox.showwarning("Import Results", 
                                         f"Imported {imported_count} employees with {error_count} errors.\n\n"
                                         f"Errors:\n{chr(10).join(error_messages)}")
                else:
                    messagebox.showinfo("Import Success", f"Successfully imported {imported_count} employees.")
                
                # Reload data
                self.load_employees()
                
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import employees: {str(e)}")

class ProjectAllocationTab(ttk.Frame):
    _MONTH_KEYS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")

    def __init__(self, parent):
        super().__init__(parent)
        self._alloc_db_id = None
        self._detail_canvas = None
        self._suppress_tree_select = False
        self._alloc_spinboxes = {}

        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Add Allocation", command=self.add_allocation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Allocation", command=self.edit_allocation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Allocation", command=self.delete_allocation).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_allocations).pack(side=tk.LEFT, padx=2)

        self.paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=2)

        self.tree_frame = ttk.Frame(left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree_yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=(
                "manager_code",
                "year",
                "cost_center",
                "work_code",
                "jan",
                "feb",
                "mar",
                "apr",
                "may",
                "jun",
                "jul",
                "aug",
                "sep",
                "oct",
                "nov",
                "dec",
            ),
            show="headings",
            yscrollcommand=scrollbar.set,
        )

        self.tree.heading("manager_code", text="Manager")
        self.tree.heading("year", text="Year")
        self.tree.heading("cost_center", text="Cost Center")
        self.tree.heading("work_code", text="Work Code")
        for m in self._MONTH_KEYS:
            self.tree.heading(m, text=m[:3].title())

        self.tree.column("manager_code", width=90)
        self.tree.column("year", width=50)
        self.tree.column("cost_center", width=80)
        self.tree.column("work_code", width=80)
        for m in self._MONTH_KEYS:
            self.tree.column(m, width=44)

        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        right_outer = ttk.Frame(self.paned, padding=(8, 0, 0, 0))
        self.paned.add(right_outer, weight=1)

        self.detail_title_var = tk.StringVar(value="Allocation")
        self._detail_canvas, form = master_detail_scroll_setup(right_outer, self.detail_title_var, "Allocation")

        form.columnconfigure(1, weight=1)
        info_frame = ttk.LabelFrame(form, text="Allocation info", padding=8)
        info_frame.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        info_frame.columnconfigure(1, weight=1)

        ttk.Label(info_frame, text="Manager Code:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.manager_var = tk.StringVar()
        self.manager_combo = ttk.Combobox(info_frame, textvariable=self.manager_var, width=18, state="readonly")
        self.manager_combo.grid(row=0, column=1, sticky=tk.W, pady=4)

        ttk.Label(info_frame, text="Year:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        years = [str(y) for y in range(datetime.now().year - 2, datetime.now().year + 3)]
        ttk.Combobox(info_frame, textvariable=self.year_var, values=years, width=8, state="readonly").grid(
            row=1, column=1, sticky=tk.W, pady=4
        )

        project_frame = ttk.LabelFrame(form, text="Project details", padding=8)
        project_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        project_frame.columnconfigure(1, weight=1)

        ttk.Label(project_frame, text="Cost Center:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(project_frame, textvariable=self.cost_center_var, width=22).grid(row=0, column=1, sticky=tk.EW, pady=4)

        ttk.Label(project_frame, text="Work Code:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.work_code_var = tk.StringVar()
        ttk.Entry(project_frame, textvariable=self.work_code_var, width=22).grid(row=1, column=1, sticky=tk.EW, pady=4)

        allocation_frame = ttk.LabelFrame(form, text="Monthly allocations", padding=8)
        allocation_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(0, 4))

        spinbox_frame = ttk.Frame(allocation_frame)
        spinbox_frame.pack(fill=tk.X, pady=4)

        month_labels = [
            ("January", "jan"),
            ("February", "feb"),
            ("March", "mar"),
            ("April", "apr"),
            ("May", "may"),
            ("June", "jun"),
            ("July", "jul"),
            ("August", "aug"),
            ("September", "sep"),
            ("October", "oct"),
            ("November", "nov"),
            ("December", "dec"),
        ]
        for i, (month_name, month_code) in enumerate(month_labels):
            row, col = i // 3, i % 3
            mf = ttk.Frame(spinbox_frame)
            mf.grid(row=row, column=col, padx=4, pady=4, sticky=tk.W)
            ttk.Label(mf, text=f"{month_name}:").pack(side=tk.LEFT)
            sb = ttk.Spinbox(mf, from_=0, to=100, increment=0.5, width=8)
            sb.pack(side=tk.LEFT, padx=4)
            self._alloc_spinboxes[month_code] = sb

        paste_frame = ttk.Frame(allocation_frame)
        paste_frame.pack(fill=tk.X, pady=6)
        ttk.Button(paste_frame, text="Paste Monthly Values", command=self._paste_monthly_values).pack(side=tk.LEFT, padx=2)

        button_bar = ttk.Frame(right_outer)
        button_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(10, 0))
        ttk.Button(button_bar, text="Save", command=self._save_detail, width=10).pack(side=tk.RIGHT, padx=4)
        ttk.Button(button_bar, text="Cancel", command=self._cancel_detail, width=10).pack(side=tk.RIGHT, padx=4)

        self.bind("<Destroy>", self._pa_on_destroy)
        self.after_idle(self._pa_set_initial_sash)
        self._refresh_manager_list()
        self.load_allocations()
        self._clear_detail_form(new_mode=True)

    def tree_yview(self, *args):
        self.tree.yview(*args)

    def _pa_set_initial_sash(self):
        try:
            self.paned.sashpos(0, 560)
        except tk.TclError:
            pass

    def _pa_on_destroy(self, event):
        if event.widget is not self:
            return
        try:
            if self._detail_canvas is not None:
                self._detail_canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass

    def _refresh_manager_list(self):
        session = get_session()
        try:
            managers = session.query(Employee.manager_code).distinct().all()
            self.manager_combo["values"] = [m[0] for m in managers if m[0]]
        finally:
            session.close()

    def _clear_detail_form(self, new_mode=False):
        self._alloc_db_id = None
        self.detail_title_var.set("New allocation" if new_mode else "Allocation")
        self.manager_var.set("")
        self.year_var.set(str(datetime.now().year))
        self.cost_center_var.set("")
        self.work_code_var.set("")
        for mk in self._MONTH_KEYS:
            self._alloc_spinboxes[mk].set(0)

    def _apply_allocation_to_form(self, allocation):
        self._alloc_db_id = allocation.id
        self.detail_title_var.set(f"Allocation #{allocation.id}")
        self.manager_var.set(allocation.manager_code or "")
        self.year_var.set(str(allocation.year))
        self.cost_center_var.set(allocation.cost_center or "")
        self.work_code_var.set(allocation.work_code or "")
        self._alloc_spinboxes["jan"].set(allocation.jan or 0)
        self._alloc_spinboxes["feb"].set(allocation.feb or 0)
        self._alloc_spinboxes["mar"].set(allocation.mar or 0)
        self._alloc_spinboxes["apr"].set(allocation.apr or 0)
        self._alloc_spinboxes["may"].set(allocation.may or 0)
        self._alloc_spinboxes["jun"].set(allocation.jun or 0)
        self._alloc_spinboxes["jul"].set(allocation.jul or 0)
        self._alloc_spinboxes["aug"].set(allocation.aug or 0)
        self._alloc_spinboxes["sep"].set(allocation.sep or 0)
        self._alloc_spinboxes["oct"].set(allocation.oct or 0)
        self._alloc_spinboxes["nov"].set(allocation.nov or 0)
        self._alloc_spinboxes["dec"].set(allocation.dec or 0)

    def _load_allocation_by_key(self, manager_code, year, cost_center, work_code):
        session = get_session()
        try:
            allocation = (
                session.query(ProjectAllocation)
                .filter(
                    ProjectAllocation.manager_code == manager_code,
                    ProjectAllocation.year == year,
                    ProjectAllocation.cost_center == cost_center,
                    ProjectAllocation.work_code == work_code,
                )
                .first()
            )
            if allocation:
                self._apply_allocation_to_form(allocation)
            else:
                messagebox.showerror("Error", "Allocation not found.")
                self._clear_detail_form(new_mode=True)
        finally:
            session.close()

    def _on_tree_select(self, _event=None):
        if self._suppress_tree_select:
            return
        selected = self.tree.selection()
        if not selected:
            return
        v = self.tree.item(selected[0])["values"]
        self._load_allocation_by_key(v[0], v[1], v[2], v[3])

    def _collect_detail_payload(self):
        if not self.manager_var.get():
            raise ValueError("Manager Code is required")
        if not self.year_var.get():
            raise ValueError("Year is required")
        if not self.cost_center_var.get().strip():
            raise ValueError("Cost Center is required")
        if not self.work_code_var.get().strip():
            raise ValueError("Work Code is required")
        monthly = {}
        for month, spinbox in self._alloc_spinboxes.items():
            try:
                monthly[month] = float(spinbox.get())
            except ValueError:
                raise ValueError(f"Invalid value for {month}")
        return {
            "manager_code": self.manager_var.get(),
            "year": int(self.year_var.get()),
            "cost_center": self.cost_center_var.get().strip(),
            "work_code": self.work_code_var.get().strip(),
            **monthly,
        }

    def _paste_monthly_values(self):
        try:
            clipboard = self.clipboard_get()
            raw = clipboard.strip().split()
            float_values = []
            for val in raw:
                val = val.replace(",", "").strip()
                float_val = float(val)
                if float_val < 0 or float_val > 100:
                    raise ValueError(f"Value {float_val} is out of range (0-100)")
                float_values.append(float_val)
            if len(float_values) != 12:
                messagebox.showerror(
                    "Error",
                    f"Expected 12 values, got {len(float_values)}.\nPlease copy exactly 12 monthly values.",
                )
                return
            for month, value in zip(self._MONTH_KEYS, float_values):
                self._alloc_spinboxes[month].set(f"{value:.1f}")
            messagebox.showinfo("Success", "Values pasted successfully.")
        except tk.TclError:
            messagebox.showerror("Error", "No data in clipboard. Please copy values first.")
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def _save_detail(self):
        try:
            data = self._collect_detail_payload()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        try:
            session = get_session()
            if self._alloc_db_id is None:
                allocation = ProjectAllocation(
                    manager_code=data["manager_code"],
                    year=data["year"],
                    cost_center=data["cost_center"],
                    work_code=data["work_code"],
                    jan=data["jan"],
                    feb=data["feb"],
                    mar=data["mar"],
                    apr=data["apr"],
                    may=data["may"],
                    jun=data["jun"],
                    jul=data["jul"],
                    aug=data["aug"],
                    sep=data["sep"],
                    oct=data["oct"],
                    nov=data["nov"],
                    dec=data["dec"],
                )
                session.add(allocation)
                session.commit()
                new_id = allocation.id
                session.close()
                self._refresh_manager_list()
                self.load_allocations()
                self._select_tree_row_by_db_id(new_id)
                messagebox.showinfo("Success", "Allocation added successfully.")
            else:
                allocation = session.query(ProjectAllocation).filter(ProjectAllocation.id == self._alloc_db_id).first()
                if not allocation:
                    session.close()
                    messagebox.showerror("Error", "Allocation not found.")
                    return
                allocation.manager_code = data["manager_code"]
                allocation.year = data["year"]
                allocation.cost_center = data["cost_center"]
                allocation.work_code = data["work_code"]
                allocation.jan = data["jan"]
                allocation.feb = data["feb"]
                allocation.mar = data["mar"]
                allocation.apr = data["apr"]
                allocation.may = data["may"]
                allocation.jun = data["jun"]
                allocation.jul = data["jul"]
                allocation.aug = data["aug"]
                allocation.sep = data["sep"]
                allocation.oct = data["oct"]
                allocation.nov = data["nov"]
                allocation.dec = data["dec"]
                session.commit()
                aid = self._alloc_db_id
                session.close()
                self._refresh_manager_list()
                self.load_allocations()
                self._select_tree_row_by_db_id(aid)
                messagebox.showinfo("Success", "Allocation updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save allocation: {str(e)}")

    def _select_tree_row_by_db_id(self, alloc_id):
        self._suppress_tree_select = True
        try:
            session = get_session()
            try:
                allocation = session.query(ProjectAllocation).filter(ProjectAllocation.id == alloc_id).first()
            finally:
                session.close()
            if not allocation:
                return
            key = (allocation.manager_code, allocation.year, allocation.cost_center, allocation.work_code)
            for item in self.tree.get_children():
                v = self.tree.item(item)["values"]
                if (v[0], v[1], v[2], v[3]) == key:
                    self.tree.selection_set(item)
                    self.tree.see(item)
                    self._load_allocation_by_key(key[0], key[1], key[2], key[3])
                    return
        finally:
            self._suppress_tree_select = False

    def _cancel_detail(self):
        if self._alloc_db_id is not None:
            session = get_session()
            try:
                allocation = session.query(ProjectAllocation).filter(ProjectAllocation.id == self._alloc_db_id).first()
                if allocation:
                    self._apply_allocation_to_form(allocation)
            finally:
                session.close()
        else:
            sel = self.tree.selection()
            if sel:
                self._on_tree_select()
            else:
                self._clear_detail_form(new_mode=True)

    def load_allocations(self):
        """Load project allocations from database"""
        keep_id = self._alloc_db_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            session = get_session()
            allocations = session.query(ProjectAllocation).all()
            count = 0
            for allocation in allocations:
                item_id = self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        allocation.manager_code,
                        allocation.year,
                        allocation.cost_center,
                        allocation.work_code,
                        allocation.jan,
                        allocation.feb,
                        allocation.mar,
                        allocation.apr,
                        allocation.may,
                        allocation.jun,
                        allocation.jul,
                        allocation.aug,
                        allocation.sep,
                        allocation.oct,
                        allocation.nov,
                        allocation.dec,
                    ),
                )
                self.tree.item(item_id, tags=("evenrow",) if count % 2 == 1 else ("oddrow",))
                count += 1
            self.tree.tag_configure("oddrow", background=COLORS["white"])
            self.tree.tag_configure("evenrow", background=COLORS["table_row_alt"])
            session.close()
            if keep_id is not None:
                if self._tree_has_db_row(keep_id):
                    self._select_tree_row_by_db_id(keep_id)
                else:
                    self._clear_detail_form(new_mode=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load allocations: {str(e)}")

    def _tree_has_db_row(self, alloc_id):
        session = get_session()
        try:
            return session.query(ProjectAllocation).filter(ProjectAllocation.id == alloc_id).first() is not None
        finally:
            session.close()

    def add_allocation(self):
        self._suppress_tree_select = True
        try:
            self.tree.selection_remove(self.tree.selection())
        except tk.TclError:
            pass
        self._suppress_tree_select = False
        self._refresh_manager_list()
        self._clear_detail_form(new_mode=True)

    def edit_allocation(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an allocation to edit.")
            return
        v = self.tree.item(selected[0])["values"]
        self._load_allocation_by_key(v[0], v[1], v[2], v[3])

    def delete_allocation(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an allocation to delete.")
            return
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this allocation?"):
            return
        v = self.tree.item(selected[0])["values"]
        try:
            session = get_session()
            allocation = (
                session.query(ProjectAllocation)
                .filter(
                    ProjectAllocation.manager_code == v[0],
                    ProjectAllocation.year == v[1],
                    ProjectAllocation.cost_center == v[2],
                    ProjectAllocation.work_code == v[3],
                )
                .first()
            )
            if allocation:
                deleted_id = allocation.id
                session.delete(allocation)
                session.commit()
                if self._alloc_db_id == deleted_id:
                    self._clear_detail_form(new_mode=True)
                self.load_allocations()
                messagebox.showinfo("Success", "Allocation deleted successfully.")
            else:
                messagebox.showerror("Error", "Allocation not found.")
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete allocation: {str(e)}")


class ForecastVisualization(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.create_widgets()
    
    def create_widgets(self):
        # Create main container with padding
        main_frame = ttk.Frame(self, style='Card.TFrame', padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame,
            text="Forecast Reports & Analytics",
            font=('Helvetica', 14, 'bold'),
            foreground=COLORS['primary']
        )
        title_label.pack(side=tk.LEFT)
        
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=(0, 20))
        
        # Control frame for options
        control_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Year selection
        year_frame = ttk.Frame(control_frame)
        year_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(
            year_frame,
            text="Year:",
            foreground=COLORS['text']
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        self.year_combo = ttk.Combobox(
            year_frame,
            textvariable=self.year_var,
            values=[str(year) for year in range(2020, 2031)],
            width=6,
            state='readonly'
        )
        self.year_combo.pack(side=tk.LEFT)
        
        # Chart type selection
        chart_frame = ttk.Frame(control_frame)
        chart_frame.pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(
            chart_frame,
            text="Chart Type:",
            foreground=COLORS['text']
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        self.chart_type_var = tk.StringVar(value="Monthly Forecast")
        self.chart_type_combo = ttk.Combobox(
            chart_frame,
            textvariable=self.chart_type_var,
            values=["Monthly Forecast", "Manager Allocation", "Employee Type Distribution", "GA01 Weeks", "Planned Changes"],
            width=20,
            state='readonly'
        )
        self.chart_type_combo.pack(side=tk.LEFT)
        
        # Generate button
        ttk.Button(
            control_frame,
            text="Generate Chart",
            style='Primary.TButton',
            command=self.generate_chart
        ).pack(side=tk.LEFT)
        
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=20)
        
        # Chart container
        self.chart_frame = ttk.Frame(main_frame, style='Card.TFrame', padding="10")
        self.chart_frame.pack(fill=tk.BOTH, expand=True)
        
        # Bind events
        self.year_combo.bind('<<ComboboxSelected>>', lambda e: self.generate_chart())
        self.chart_type_combo.bind('<<ComboboxSelected>>', lambda e: self.generate_chart())
    
    def generate_chart(self):
        # Clear previous chart
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        
        # Create figure with white background
        fig = Figure(figsize=(10, 6), facecolor=COLORS['white'])
        ax = fig.add_subplot(111)
        
        # Set background color
        ax.set_facecolor(COLORS['white'])
        
        # Get selected year
        year = int(self.year_var.get())
        
        # Generate selected chart type
        chart_type = self.chart_type_var.get()
        if chart_type == "Monthly Forecast":
            self._generate_monthly_forecast_chart(ax, year)
        elif chart_type == "Manager Allocation":
            self._generate_manager_allocation_chart(ax, year)
        elif chart_type == "Employee Type Distribution":
            self._generate_employee_type_distribution(ax, year)
        elif chart_type == "GA01 Weeks":
            self._generate_ga01_weeks_chart(ax, year)
        elif chart_type == "Planned Changes":
            self._generate_planned_changes_chart(ax, year)
        
        # Create canvas
        canvas = FigureCanvasTkAgg(fig, master=self.chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def _generate_monthly_forecast_chart(self, ax, year):
        session = get_session()
        forecasts = session.query(Forecast).filter(Forecast.year == year).all()
        session.close()
        
        if not forecasts:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        total_hours = [0] * 12
        
        for forecast in forecasts:
            total_hours[0] += forecast.jan
            total_hours[1] += forecast.feb
            total_hours[2] += forecast.mar
            total_hours[3] += forecast.apr
            total_hours[4] += forecast.may
            total_hours[5] += forecast.jun
            total_hours[6] += forecast.jul
            total_hours[7] += forecast.aug
            total_hours[8] += forecast.sep
            total_hours[9] += forecast.oct
            total_hours[10] += forecast.nov
            total_hours[11] += forecast.dec
        
        # Create bar chart
        bars = ax.bar(months, total_hours, color=COLORS['secondary'])
        
        # Customize chart
        ax.set_title(f'Monthly Forecast Hours - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        ax.set_xlabel('Month', color=COLORS['text'])
        ax.set_ylabel('Total Hours', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_manager_allocation_chart(self, ax, year):
        session = get_session()
        forecasts = session.query(Forecast).filter(Forecast.year == year).all()
        session.close()
        
        if not forecasts:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        manager_data = {}
        for forecast in forecasts:
            if forecast.manager_code not in manager_data:
                manager_data[forecast.manager_code] = 0
            manager_data[forecast.manager_code] += forecast.total_hours
        
        # Sort managers by total hours
        sorted_managers = sorted(manager_data.items(), 
                               key=lambda x: x[1], 
                               reverse=True)
        
        # Create horizontal bar chart
        managers = [m[0] for m in sorted_managers]
        hours = [m[1] for m in sorted_managers]
        
        bars = ax.barh(managers, hours, color=COLORS['secondary'])
        
        # Customize chart
        ax.set_title(f'Manager Allocation - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        ax.set_xlabel('Total Hours', color=COLORS['text'])
        ax.set_ylabel('Manager Code', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2.,
                   f'{int(width)}',
                   ha='left', va='center', pad=5)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_employee_type_distribution(self, ax, year):
        session = get_session()
        employees = session.query(Employee).all()
        session.close()
        
        if not employees:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Count employee types
        type_counts = {}
        for employee in employees:
            if employee.employment_type not in type_counts:
                type_counts[employee.employment_type] = 0
            type_counts[employee.employment_type] += 1
        
        # Create pie chart
        types = list(type_counts.keys())
        counts = list(type_counts.values())
        
        colors = [COLORS['secondary'], COLORS['accent']]
        patches, texts, autotexts = ax.pie(counts, 
                                         labels=types,
                                         colors=colors,
                                         autopct='%1.1f%%',
                                         startangle=90)
        
        # Customize chart
        ax.set_title(f'Employee Type Distribution - {year}', 
                    color=COLORS['primary'], 
                    pad=20)
        
        # Customize text properties
        plt.setp(autotexts, size=9, weight="bold")
        plt.setp(texts, size=10)
    
    def _generate_ga01_weeks_chart(self, ax, year):
        session = get_session()
        ga01_weeks = session.query(GA01Week).filter(GA01Week.year == year).all()
        session.close()
        
        if not ga01_weeks:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        weeks = [week.weeks for week in ga01_weeks]
        
        # Create bar chart
        bars = ax.bar(months, weeks, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'GA01 Weeks - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Month', color=COLORS['text'])
        ax.set_ylabel('Weeks', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Rotate x-axis labels for better readability
        plt.xticks(rotation=45)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_planned_changes_chart(self, ax, year):
        session = get_session()
        changes = session.query(PlannedChange).filter(PlannedChange.effective_date.between(
            datetime(year, 1, 1).date(), datetime(year, 12, 31).date())).all()
        session.close()
        
        if not changes:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        change_types = ['New Hire', 'Conversion', 'Termination']
        change_counts = [sum(1 for change in changes if change.change_type == change_type) for change_type in change_types]
        
        # Create bar chart
        bars = ax.bar(change_types, change_counts, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'Planned Changes - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Change Type', color=COLORS['text'])
        ax.set_ylabel('Count', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}',
                   ha='center', va='bottom')
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)
    
    def _generate_allocation_chart(self, ax, year):
        session = get_session()
        allocations = session.query(ProjectAllocation).filter(ProjectAllocation.year == year).all()
        session.close()
        
        if not allocations:
            ax.text(0.5, 0.5, 'No data available', 
                   horizontalalignment='center',
                   verticalalignment='center',
                   transform=ax.transAxes)
            return
        
        # Prepare data
        manager_data = {}
        for allocation in allocations:
            if allocation.manager_code not in manager_data:
                manager_data[allocation.manager_code] = 0
            manager_data[allocation.manager_code] += allocation.total_hours
        
        # Sort managers by total hours
        sorted_managers = sorted(manager_data.items(), 
                               key=lambda x: x[1], 
                               reverse=True)
        
        # Create horizontal bar chart
        managers = [m[0] for m in sorted_managers]
        hours = [m[1] for m in sorted_managers]
        
        bars = ax.barh(managers, hours, color=COLORS['accent'])
        
        # Customize chart
        ax.set_title(f'Project Allocations - {year}', 
                    color=COLORS['accent'], 
                    pad=20)
        ax.set_xlabel('Total Hours', color=COLORS['text'])
        ax.set_ylabel('Manager Code', color=COLORS['text'])
        
        # Add value labels on bars
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2.,
                   f'{int(width)}',
                   ha='left', va='center', pad=5)
        
        # Add grid
        ax.grid(True, linestyle='--', alpha=0.7)

class ForecastTab(ttk.Frame):
    _MONTH_KEYS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")

    def __init__(self, parent):
        super().__init__(parent)
        self._forecast_id = None
        self._detail_canvas = None
        self._suppress_tree_select = False
        self._fc_spinboxes = {}

        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Calculate Forecast", command=self.calculate_forecast, style="Primary.TButton").pack(
            side=tk.LEFT, padx=2
        )
        ttk.Button(toolbar, text="Add Forecast", command=self.add_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Forecast", command=self.edit_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Forecast", command=self.delete_forecast).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_forecasts).pack(side=tk.LEFT, padx=2)

        year_frame = ttk.Frame(toolbar)
        year_frame.pack(side=tk.RIGHT, padx=5)

        ttk.Label(year_frame, text="Year:").pack(side=tk.LEFT, padx=2)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        year_combo = ttk.Combobox(
            year_frame,
            textvariable=self.year_var,
            values=[str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)],
            width=6,
            state="readonly",
        )
        year_combo.pack(side=tk.LEFT, padx=2)
        year_combo.bind("<<ComboboxSelected>>", lambda e: self.load_forecasts())

        self.paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=2)

        self.tree_frame = ttk.Frame(left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree_yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=(
                "id",
                "manager_code",
                "cost_center",
                "work_code",
                "jan",
                "feb",
                "mar",
                "apr",
                "may",
                "jun",
                "jul",
                "aug",
                "sep",
                "oct",
                "nov",
                "dec",
                "total",
            ),
            show="headings",
            yscrollcommand=scrollbar.set,
        )

        self.tree.heading("id", text="ID")
        self.tree.heading("manager_code", text="Manager")
        self.tree.heading("cost_center", text="Cost Center")
        self.tree.heading("work_code", text="Work Code")
        for m in self._MONTH_KEYS:
            self.tree.heading(m, text=m[:3].title())
        self.tree.heading("total", text="Total")

        self.tree.column("id", width=50)
        self.tree.column("manager_code", width=72)
        self.tree.column("cost_center", width=72)
        self.tree.column("work_code", width=72)
        for m in self._MONTH_KEYS:
            self.tree.column(m, width=44)
        self.tree.column("total", width=64)

        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        right_outer = ttk.Frame(self.paned, padding=(8, 0, 0, 0))
        self.paned.add(right_outer, weight=1)

        self.detail_title_var = tk.StringVar(value="Forecast")
        self._detail_canvas, form = master_detail_scroll_setup(right_outer, self.detail_title_var, "Forecast")

        form.columnconfigure(1, weight=1)
        info_frame = ttk.LabelFrame(form, text="Forecast info", padding=8)
        info_frame.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))
        info_frame.columnconfigure(1, weight=1)

        ttk.Label(info_frame, text="Manager Code:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.manager_var = tk.StringVar()
        self.managers_combo = ttk.Combobox(info_frame, textvariable=self.manager_var, width=20)
        self.managers_combo.grid(row=0, column=1, sticky=tk.W, pady=4)

        ttk.Label(info_frame, text="Year:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.forecast_year_var = tk.StringVar(value=str(datetime.now().year))
        fy = [str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)]
        ttk.Combobox(info_frame, textvariable=self.forecast_year_var, values=fy, width=8).grid(
            row=1, column=1, sticky=tk.W, pady=4
        )

        ttk.Label(info_frame, text="Cost Center:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.cost_center_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.cost_center_var, width=22).grid(row=2, column=1, sticky=tk.EW, pady=4)

        ttk.Label(info_frame, text="Work Code:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.work_code_var = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.work_code_var, width=22).grid(row=3, column=1, sticky=tk.EW, pady=4)

        hours_frame = ttk.LabelFrame(form, text="Monthly hours", padding=8)
        hours_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 4))

        spinbox_frame = ttk.Frame(hours_frame)
        spinbox_frame.pack(fill=tk.X, pady=4)

        month_labels = [
            ("January", "jan"),
            ("February", "feb"),
            ("March", "mar"),
            ("April", "apr"),
            ("May", "may"),
            ("June", "jun"),
            ("July", "jul"),
            ("August", "aug"),
            ("September", "sep"),
            ("October", "oct"),
            ("November", "nov"),
            ("December", "dec"),
        ]
        for i, (month_name, month_code) in enumerate(month_labels):
            row, col = i // 3, i % 3
            mf = ttk.Frame(spinbox_frame)
            mf.grid(row=row, column=col, padx=4, pady=4, sticky=tk.W)
            ttk.Label(mf, text=f"{month_name}:").pack(side=tk.LEFT)
            sb = ttk.Spinbox(mf, from_=0, to=1000, increment=0.5, width=8)
            sb.pack(side=tk.LEFT, padx=4)
            self._fc_spinboxes[month_code] = sb

        button_bar = ttk.Frame(right_outer)
        button_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(10, 0))
        ttk.Button(button_bar, text="Save", command=self._save_detail, width=10).pack(side=tk.RIGHT, padx=4)
        ttk.Button(button_bar, text="Cancel", command=self._cancel_detail, width=10).pack(side=tk.RIGHT, padx=4)

        self.bind("<Destroy>", self._fc_on_destroy)
        self.after_idle(self._fc_set_initial_sash)
        self._refresh_manager_list()
        self.load_forecasts()
        self._clear_detail_form(new_mode=True)

    def tree_yview(self, *args):
        self.tree.yview(*args)

    def _fc_set_initial_sash(self):
        try:
            self.paned.sashpos(0, 560)
        except tk.TclError:
            pass

    def _fc_on_destroy(self, event):
        if event.widget is not self:
            return
        try:
            if self._detail_canvas is not None:
                self._detail_canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass

    def _refresh_manager_list(self):
        session = get_session()
        try:
            managers = session.query(Employee.manager_code).distinct().all()
            self.managers_combo["values"] = [m[0] for m in managers if m[0]]
        finally:
            session.close()

    def _clear_detail_form(self, new_mode=False):
        self._forecast_id = None
        self.detail_title_var.set("New forecast" if new_mode else "Forecast")
        self.manager_var.set("")
        self.forecast_year_var.set(self.year_var.get())
        self.cost_center_var.set("")
        self.work_code_var.set("")
        for mk in self._MONTH_KEYS:
            self._fc_spinboxes[mk].set(0)

    def _apply_forecast_to_form(self, forecast):
        self._forecast_id = forecast.id
        self.detail_title_var.set(f"Forecast #{forecast.id}")
        self.manager_var.set(forecast.manager_code or "")
        self.forecast_year_var.set(str(forecast.year))
        self.cost_center_var.set(forecast.cost_center or "")
        self.work_code_var.set(forecast.work_code or "")
        self._fc_spinboxes["jan"].set(forecast.jan or 0)
        self._fc_spinboxes["feb"].set(forecast.feb or 0)
        self._fc_spinboxes["mar"].set(forecast.mar or 0)
        self._fc_spinboxes["apr"].set(forecast.apr or 0)
        self._fc_spinboxes["may"].set(forecast.may or 0)
        self._fc_spinboxes["jun"].set(forecast.jun or 0)
        self._fc_spinboxes["jul"].set(forecast.jul or 0)
        self._fc_spinboxes["aug"].set(forecast.aug or 0)
        self._fc_spinboxes["sep"].set(forecast.sep or 0)
        self._fc_spinboxes["oct"].set(forecast.oct or 0)
        self._fc_spinboxes["nov"].set(forecast.nov or 0)
        self._fc_spinboxes["dec"].set(forecast.dec or 0)

    def _load_forecast_by_id(self, forecast_id):
        session = get_session()
        try:
            forecast = session.query(Forecast).filter(Forecast.id == forecast_id).first()
            if forecast:
                self._apply_forecast_to_form(forecast)
            else:
                messagebox.showerror("Error", "Forecast not found.")
                self._clear_detail_form(new_mode=True)
        finally:
            session.close()

    def _on_tree_select(self, _event=None):
        if self._suppress_tree_select:
            return
        selected = self.tree.selection()
        if not selected:
            return
        forecast_id = self.tree.item(selected[0])["values"][0]
        self._load_forecast_by_id(forecast_id)

    def _collect_detail_payload(self):
        if not self.manager_var.get():
            raise ValueError("Manager Code is required")
        if not self.forecast_year_var.get():
            raise ValueError("Year is required")
        if not self.cost_center_var.get().strip():
            raise ValueError("Cost Center is required")
        if not self.work_code_var.get().strip():
            raise ValueError("Work Code is required")
        monthly = {}
        for month, spinbox in self._fc_spinboxes.items():
            try:
                monthly[month] = float(spinbox.get())
            except ValueError:
                raise ValueError(f"Invalid value for {month}")
        return {
            "manager_code": self.manager_var.get(),
            "year": int(self.forecast_year_var.get()),
            "cost_center": self.cost_center_var.get().strip(),
            "work_code": self.work_code_var.get().strip(),
            **monthly,
        }

    def _save_detail(self):
        try:
            data = self._collect_detail_payload()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        total_hours = sum(data[m] for m in self._MONTH_KEYS)
        try:
            session = get_session()
            if self._forecast_id is None:
                forecast = Forecast(
                    year=data["year"],
                    manager_code=data["manager_code"],
                    cost_center=data["cost_center"],
                    work_code=data["work_code"],
                    jan=data["jan"],
                    feb=data["feb"],
                    mar=data["mar"],
                    apr=data["apr"],
                    may=data["may"],
                    jun=data["jun"],
                    jul=data["jul"],
                    aug=data["aug"],
                    sep=data["sep"],
                    oct=data["oct"],
                    nov=data["nov"],
                    dec=data["dec"],
                    total_hours=total_hours,
                )
                session.add(forecast)
                session.commit()
                new_id = forecast.id
                session.close()
                self._refresh_manager_list()
                self.load_forecasts()
                self._select_tree_row_by_id(new_id)
                messagebox.showinfo("Success", "Forecast added successfully.")
            else:
                forecast = session.query(Forecast).filter(Forecast.id == self._forecast_id).first()
                if not forecast:
                    session.close()
                    messagebox.showerror("Error", "Forecast not found.")
                    return
                forecast.year = data["year"]
                forecast.manager_code = data["manager_code"]
                forecast.cost_center = data["cost_center"]
                forecast.work_code = data["work_code"]
                forecast.jan = data["jan"]
                forecast.feb = data["feb"]
                forecast.mar = data["mar"]
                forecast.apr = data["apr"]
                forecast.may = data["may"]
                forecast.jun = data["jun"]
                forecast.jul = data["jul"]
                forecast.aug = data["aug"]
                forecast.sep = data["sep"]
                forecast.oct = data["oct"]
                forecast.nov = data["nov"]
                forecast.dec = data["dec"]
                forecast.total_hours = total_hours
                session.commit()
                fid = self._forecast_id
                session.close()
                self._refresh_manager_list()
                self.load_forecasts()
                self._select_tree_row_by_id(fid)
                messagebox.showinfo("Success", "Forecast updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save forecast: {str(e)}")

    def _select_tree_row_by_id(self, forecast_id):
        self._suppress_tree_select = True
        try:
            for item in self.tree.get_children():
                if self.tree.item(item)["values"][0] == forecast_id:
                    self.tree.selection_set(item)
                    self.tree.see(item)
                    self._load_forecast_by_id(forecast_id)
                    return
        finally:
            self._suppress_tree_select = False

    def _cancel_detail(self):
        if self._forecast_id is not None:
            self._load_forecast_by_id(self._forecast_id)
        else:
            sel = self.tree.selection()
            if sel:
                self._on_tree_select()
            else:
                self._clear_detail_form(new_mode=True)

    def load_forecasts(self):
        """Load forecasts from database"""
        keep_id = self._forecast_id
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            session = get_session()
            year = int(self.year_var.get())
            forecasts = session.query(Forecast).filter(Forecast.year == year).all()
            count = 0
            for forecast in forecasts:
                total = sum(
                    [
                        forecast.jan,
                        forecast.feb,
                        forecast.mar,
                        forecast.apr,
                        forecast.may,
                        forecast.jun,
                        forecast.jul,
                        forecast.aug,
                        forecast.sep,
                        forecast.oct,
                        forecast.nov,
                        forecast.dec,
                    ]
                )
                item_id = self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        forecast.id,
                        forecast.manager_code,
                        forecast.cost_center,
                        forecast.work_code,
                        forecast.jan,
                        forecast.feb,
                        forecast.mar,
                        forecast.apr,
                        forecast.may,
                        forecast.jun,
                        forecast.jul,
                        forecast.aug,
                        forecast.sep,
                        forecast.oct,
                        forecast.nov,
                        forecast.dec,
                        total,
                    ),
                )
                self.tree.item(item_id, tags=("evenrow",) if count % 2 == 1 else ("oddrow",))
                count += 1
            self.tree.tag_configure("oddrow", background=COLORS["white"])
            self.tree.tag_configure("evenrow", background=COLORS["table_row_alt"])
            session.close()
            if keep_id is not None:
                if any(self.tree.item(i)["values"][0] == keep_id for i in self.tree.get_children()):
                    self._select_tree_row_by_id(keep_id)
                else:
                    self._clear_detail_form(new_mode=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load forecasts: {str(e)}")

    def add_forecast(self):
        self._suppress_tree_select = True
        try:
            self.tree.selection_remove(self.tree.selection())
        except tk.TclError:
            pass
        self._suppress_tree_select = False
        self._refresh_manager_list()
        self._clear_detail_form(new_mode=True)

    def edit_forecast(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a forecast to edit.")
            return
        forecast_id = self.tree.item(selected[0])["values"][0]
        self._load_forecast_by_id(forecast_id)

    def delete_forecast(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a forecast to delete.")
            return
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this forecast?"):
            return
        forecast_id = self.tree.item(selected[0])["values"][0]
        try:
            session = get_session()
            forecast = session.query(Forecast).filter(Forecast.id == forecast_id).first()
            if forecast:
                session.delete(forecast)
                session.commit()
                if self._forecast_id == forecast_id:
                    self._clear_detail_form(new_mode=True)
                self.load_forecasts()
                messagebox.showinfo("Success", "Forecast deleted successfully.")
            else:
                messagebox.showwarning("Warning", "Selected forecast not found.")
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete forecast: {str(e)}")

    def calculate_forecast(self):
        """Automatically calculate forecast based on employee data and allocations"""
        try:
            year = int(self.year_var.get())
            
            # Get confirmation before proceeding
            if not messagebox.askyesno("Confirm Calculate", f"This will calculate forecasts for {year} based on current employee data and allocations. Continue?"):
                return
            
            session = get_session()
            
            # Get settings for hours calculation
            settings = session.query(Settings).first()
            if not settings:
                messagebox.showerror("Error", "Settings not found. Please configure settings first.")
                session.close()
                return
            
            # Get all employees
            employees = session.query(Employee).all()
            if not employees:
                messagebox.showwarning("Warning", "No employees found. Please add employees first.")
                session.close()
                return
            
            # Get all allocations for the year
            allocations = session.query(ProjectAllocation).filter(ProjectAllocation.year == year).all()
            
            # Get all existing forecasts for the year to potentially update
            existing_forecasts = {}
            for forecast in session.query(Forecast).filter(Forecast.year == year).all():
                key = (forecast.manager_code, forecast.cost_center, forecast.work_code)
                existing_forecasts[key] = forecast
            
            # Track what we've processed
            processed_count = 0
            created_count = 0
            updated_count = 0
            
            # Process each employee
            for employee in employees:
                # Skip employees who ended before this year or started after this year
                if employee.end_date and employee.end_date.year < year:
                    continue
                
                # Determine work code - use employee's work code if available, otherwise use default
                work_code = employee.work_code if employee.work_code else "DEFAULT"
                
                # Determine hours based on employment type
                monthly_hours = settings.fte_hours if employee.employment_type == EmploymentType.FTE.value else settings.contractor_hours
                
                # Calculate which months the employee is active
                active_months = {month: 1.0 for month in range(1, 13)}  # Default to full month
                
                # Adjust for start date if in this year
                if employee.start_date.year == year:
                    for month in range(1, employee.start_date.month):
                        active_months[month] = 0.0
                    
                    # If starting mid-month, prorate
                    if employee.start_date.day > 1:
                        days_in_month = 30  # Approximation
                        active_months[employee.start_date.month] = (days_in_month - employee.start_date.day + 1) / days_in_month
                
                # Adjust for end date if in this year
                if employee.end_date and employee.end_date.year == year:
                    for month in range(employee.end_date.month + 1, 13):
                        active_months[month] = 0.0
                    
                    # If ending mid-month, prorate
                    if employee.end_date.day < 30:  # Approximation
                        days_in_month = 30  # Approximation
                        active_months[employee.end_date.month] = employee.end_date.day / days_in_month
                
                # Create or update forecast
                key = (employee.manager_code, employee.cost_center, work_code)
                
                # Calculate monthly hours
                month_hours = {
                    "jan": monthly_hours * active_months[1],
                    "feb": monthly_hours * active_months[2],
                    "mar": monthly_hours * active_months[3],
                    "apr": monthly_hours * active_months[4],
                    "may": monthly_hours * active_months[5],
                    "jun": monthly_hours * active_months[6],
                    "jul": monthly_hours * active_months[7],
                    "aug": monthly_hours * active_months[8],
                    "sep": monthly_hours * active_months[9],
                    "oct": monthly_hours * active_months[10],
                    "nov": monthly_hours * active_months[11],
                    "dec": monthly_hours * active_months[12]
                }
                
                # Apply any specific allocations for this employee
                # (This would require linking allocations to employees, which isn't in the current model)
                # For now, we'll just use the basic calculation
                
                # Calculate total hours
                total_hours = sum(month_hours.values())
                
                if key in existing_forecasts:
                    # Update existing forecast
                    forecast = existing_forecasts[key]
                    forecast.jan = month_hours["jan"]
                    forecast.feb = month_hours["feb"]
                    forecast.mar = month_hours["mar"]
                    forecast.apr = month_hours["apr"]
                    forecast.may = month_hours["may"]
                    forecast.jun = month_hours["jun"]
                    forecast.jul = month_hours["jul"]
                    forecast.aug = month_hours["aug"]
                    forecast.sep = month_hours["sep"]
                    forecast.oct = month_hours["oct"]
                    forecast.nov = month_hours["nov"]
                    forecast.dec = month_hours["dec"]
                    forecast.total_hours = total_hours
                    updated_count += 1
                else:
                    # Create new forecast
                    forecast = Forecast(
                        year=year,
                        manager_code=employee.manager_code,
                        cost_center=employee.cost_center,
                        work_code=work_code,
                        jan=month_hours["jan"],
                        feb=month_hours["feb"],
                        mar=month_hours["mar"],
                        apr=month_hours["apr"],
                        may=month_hours["may"],
                        jun=month_hours["jun"],
                        jul=month_hours["jul"],
                        aug=month_hours["aug"],
                        sep=month_hours["sep"],
                        oct=month_hours["oct"],
                        nov=month_hours["nov"],
                        dec=month_hours["dec"],
                        total_hours=total_hours
                    )
                    session.add(forecast)
                    created_count += 1
                
                processed_count += 1
            
            # Commit changes
            session.commit()
            session.close()
            
            # Reload forecasts
            self.load_forecasts()
            
            # Show success message
            messagebox.showinfo("Forecast Calculation Complete", 
                              f"Processed {processed_count} employees.\n"
                              f"Created {created_count} new forecasts.\n"
                              f"Updated {updated_count} existing forecasts.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate forecast: {str(e)}")

class PlannedChangesTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self._record_id = None
        self._pc_employee_id = None
        self._pc_employees = {}
        self._pc_canvas = None
        self._suppress_tree_select = False
        
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Add Change", command=self.add_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Edit Change", command=self.edit_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Delete Change", command=self.delete_change).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Refresh", command=self.load_changes).pack(side=tk.LEFT, padx=2)
        
        year_frame = ttk.Frame(toolbar)
        year_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(year_frame, text="Year:").pack(side=tk.LEFT, padx=2)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        year_combo = ttk.Combobox(
            year_frame,
            textvariable=self.year_var,
            values=[str(y) for y in range(datetime.now().year - 2, datetime.now().year + 5)],
            width=6,
            state="readonly"
        )
        year_combo.pack(side=tk.LEFT, padx=2)
        year_combo.bind("<<ComboboxSelected>>", lambda e: self.load_changes())
        
        self.paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=2)
        
        self.tree_frame = ttk.Frame(left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree_yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=(
                "id", "description", "change_type", "effective_date", "name", "team", "manager_code", "status"
            ),
            show="headings",
            yscrollcommand=scrollbar.set,
        )
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        self.tree.heading("id", text="ID")
        self.tree.heading("description", text="Description")
        self.tree.heading("change_type", text="Change Type")
        self.tree.heading("effective_date", text="Effective Date")
        self.tree.heading("name", text="Name")
        self.tree.heading("team", text="Team")
        self.tree.heading("manager_code", text="Manager")
        self.tree.heading("status", text="Status")
        
        self.tree.column("id", width=50)
        self.tree.column("description", width=200)
        self.tree.column("change_type", width=100)
        self.tree.column("effective_date", width=100)
        self.tree.column("name", width=150)
        self.tree.column("team", width=100)
        self.tree.column("manager_code", width=100)
        self.tree.column("status", width=80)
        
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        
        right_outer = ttk.Frame(self.paned, padding=(8, 0, 0, 0))
        self.paned.add(right_outer, weight=1)
        
        self._build_detail_pane(right_outer)
        
        self.bind("<Destroy>", self._pc_on_destroy)
        
        self.after_idle(self._pc_set_initial_sash)
        self.load_changes()
        self._clear_detail_form(new_mode=True)
    
    def tree_yview(self, *args):
        self.tree.yview(*args)
    
    def _pc_set_initial_sash(self):
        try:
            self.paned.sashpos(0, 560)
        except tk.TclError:
            pass
    
    def _pc_on_destroy(self, event):
        if event.widget is not self:
            return
        try:
            if self._pc_canvas is not None:
                self._pc_canvas.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
    
    def _build_detail_pane(self, parent):
        self.detail_title_var = tk.StringVar()
        self._pc_canvas, form = master_detail_scroll_setup(parent, self.detail_title_var, "Planned change")

        form.columnconfigure(1, weight=1)
        
        ttk.Label(form, text="Description:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.description_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.description_var, width=36).grid(row=0, column=1, sticky=tk.EW, pady=4)
        
        ttk.Label(form, text="Change Type:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.change_type_var = tk.StringVar()
        change_types = [ct.value for ct in ChangeType]
        self.change_type_combo = ttk.Combobox(
            form, textvariable=self.change_type_var, values=change_types, width=18, state="readonly"
        )
        self.change_type_combo.grid(row=1, column=1, sticky=tk.W, pady=4)
        self.change_type_combo.bind("<<ComboboxSelected>>", self._pc_on_change_type_selected)
        
        ttk.Label(form, text="Effective Date (mm/dd/yy):").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.effective_date_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.effective_date_var, width=18).grid(row=2, column=1, sticky=tk.W, pady=4)
        
        self.employee_selection_frame = ttk.LabelFrame(form, text="Select Employee", padding=8)
        self.employee_selection_frame.grid(row=3, column=0, columnspan=2, pady=8, sticky=tk.EW)
        self.employee_selection_frame.grid_remove()
        
        lb_frame = ttk.Frame(self.employee_selection_frame)
        lb_frame.pack(fill=tk.BOTH, expand=True, pady=4)
        
        self.employee_listbox = tk.Listbox(lb_frame, height=5, width=34)
        self.employee_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        lb_sb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self.employee_listbox.yview)
        lb_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.employee_listbox.config(yscrollcommand=lb_sb.set)
        self.employee_listbox.bind("<<ListboxSelect>>", self._pc_on_employee_selected)
        
        self.details_frame = ttk.LabelFrame(form, text="Employee Details", padding=8)
        self.details_frame.grid(row=4, column=0, columnspan=2, pady=8, sticky=tk.EW)
        
        ttk.Label(self.details_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(self.details_frame, textvariable=self.name_var, width=28)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=4)
        
        ttk.Label(self.details_frame, text="Team:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.team_var = tk.StringVar()
        self.team_entry = ttk.Entry(self.details_frame, textvariable=self.team_var, width=28)
        self.team_entry.grid(row=1, column=1, sticky=tk.W, pady=4)
        
        ttk.Label(self.details_frame, text="Manager Code:").grid(row=2, column=0, sticky=tk.W, pady=4)
        self.manager_code_var = tk.StringVar()
        self.manager_code_entry = ttk.Entry(self.details_frame, textvariable=self.manager_code_var, width=18)
        self.manager_code_entry.grid(row=2, column=1, sticky=tk.W, pady=4)
        
        ttk.Label(self.details_frame, text="Cost Center:").grid(row=3, column=0, sticky=tk.W, pady=4)
        self.cost_center_var = tk.StringVar()
        self.cost_center_entry = ttk.Entry(self.details_frame, textvariable=self.cost_center_var, width=18)
        self.cost_center_entry.grid(row=3, column=1, sticky=tk.W, pady=4)
        
        ttk.Label(self.details_frame, text="Employment Type:").grid(row=4, column=0, sticky=tk.W, pady=4)
        self.employment_type_var = tk.StringVar()
        employment_types = [et.value for et in EmploymentType]
        self.employment_type_combo = ttk.Combobox(
            self.details_frame, textvariable=self.employment_type_var, values=employment_types, width=18
        )
        self.employment_type_combo.grid(row=4, column=1, sticky=tk.W, pady=4)
        
        ttk.Label(form, text="Status:").grid(row=5, column=0, sticky=tk.W, pady=4)
        self.status_var = tk.StringVar(value="Planned")
        statuses = ["Planned", "In Progress", "Completed", "Cancelled"]
        ttk.Combobox(form, textvariable=self.status_var, values=statuses, width=18, state="readonly").grid(
            row=5, column=1, sticky=tk.W, pady=4
        )
        
        button_bar = ttk.Frame(parent)
        button_bar.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=(10, 0))
        ttk.Button(button_bar, text="Save", command=self._save_detail, width=10).pack(side=tk.RIGHT, padx=4)
        ttk.Button(button_bar, text="Cancel", command=self._cancel_detail, width=10).pack(side=tk.RIGHT, padx=4)
    
    def _clear_detail_form(self, new_mode=False):
        self._record_id = None
        self._pc_employee_id = None
        self.detail_title_var.set("New planned change" if new_mode else "Planned change")
        self.description_var.set("")
        self.change_type_var.set("")
        self.effective_date_var.set(datetime.now().strftime("%m/%d/%y"))
        self.name_var.set("")
        self.team_var.set("")
        self.manager_code_var.set("")
        self.cost_center_var.set("")
        self.employment_type_var.set("")
        self.status_var.set("Planned")
        self.employee_selection_frame.grid_remove()
        for w in (self.name_entry, self.team_entry, self.manager_code_entry, self.cost_center_entry):
            w.config(state="normal")
        self.employment_type_combo.config(state="normal")
    
    def _apply_change_to_form(self, change):
        self._record_id = change.id
        self._pc_employee_id = change.employee_id
        self.detail_title_var.set(f"Planned change #{change.id}")
        self.description_var.set(change.description or "")
        self.change_type_var.set(change.change_type or "")
        self.effective_date_var.set(
            change.effective_date.strftime("%m/%d/%y") if change.effective_date else datetime.now().strftime("%m/%d/%y")
        )
        self.name_var.set(change.name or "")
        self.team_var.set(change.team or "")
        self.manager_code_var.set(change.manager_code or "")
        self.cost_center_var.set(change.cost_center or "")
        self.employment_type_var.set(change.employment_type or "")
        self.status_var.set(change.status or "Planned")
        
        if change.change_type in (ChangeType.CONVERSION.value, ChangeType.TERMINATION.value):
            self.employee_selection_frame.grid()
            self._pc_load_employees()
            if change.employee_id:
                self._pc_load_employee_details(change.employee_id)
        else:
            self.employee_selection_frame.grid_remove()
            for w in (self.name_entry, self.team_entry, self.manager_code_entry, self.cost_center_entry):
                w.config(state="normal")
            self.employment_type_combo.config(state="normal")
    
    def _load_change_by_id(self, change_id):
        session = get_session()
        try:
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            if change:
                self._apply_change_to_form(change)
            else:
                messagebox.showerror("Error", "Planned change not found.")
                self._clear_detail_form(new_mode=True)
        finally:
            session.close()
    
    def _on_tree_select(self, _event=None):
        if self._suppress_tree_select:
            return
        selected = self.tree.selection()
        if not selected:
            return
        change_id = self.tree.item(selected[0])["values"][0]
        self._load_change_by_id(change_id)
    
    def _pc_on_change_type_selected(self, _event=None):
        change_type = self.change_type_var.get()
        if change_type in (ChangeType.CONVERSION.value, ChangeType.TERMINATION.value):
            self.employee_selection_frame.grid()
            self._pc_load_employees()
        else:
            self.employee_selection_frame.grid_remove()
            self._pc_employee_id = None
            self.name_var.set("")
            self.team_var.set("")
            self.manager_code_var.set("")
            self.cost_center_var.set("")
            self.employment_type_var.set("")
            for w in (self.name_entry, self.team_entry, self.manager_code_entry, self.cost_center_entry):
                w.config(state="normal")
            self.employment_type_combo.config(state="normal")
    
    def _pc_load_employees(self):
        try:
            self.employee_listbox.delete(0, tk.END)
            self._pc_employees = {}
            session = get_session()
            employees = session.query(Employee).all()
            session.close()
            for emp in employees:
                self.employee_listbox.insert(tk.END, f"{emp.name} - {emp.manager_code} ({emp.employment_type})")
                idx = self.employee_listbox.size() - 1
                self._pc_employees[idx] = emp
                if self._pc_employee_id and emp.id == self._pc_employee_id:
                    self.employee_listbox.selection_set(idx)
                    self.employee_listbox.see(idx)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employees: {str(e)}")
    
    def _pc_on_employee_selected(self, _event=None):
        sel = self.employee_listbox.curselection()
        if not sel:
            return
        emp = self._pc_employees.get(sel[0])
        if emp:
            self._pc_employee_id = emp.id
            self._pc_load_employee_details(emp.id)
    
    def _pc_load_employee_details(self, employee_id):
        try:
            session = get_session()
            employee = session.query(Employee).filter(Employee.id == employee_id).first()
            session.close()
            if not employee:
                return
            self.name_var.set(employee.name)
            self.team_var.set("")
            self.manager_code_var.set(employee.manager_code)
            self.cost_center_var.set(employee.cost_center)
            self.employment_type_var.set(employee.employment_type)
            for w in (self.name_entry, self.team_entry, self.manager_code_entry, self.cost_center_entry):
                w.config(state="readonly")
            self.employment_type_combo.config(state="readonly")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load employee details: {str(e)}")
    
    def _collect_detail_payload(self):
        if not self.description_var.get().strip():
            raise ValueError("Description is required")
        if not self.change_type_var.get():
            raise ValueError("Change Type is required")
        if not self.effective_date_var.get().strip():
            raise ValueError("Effective Date is required")
        if not self.status_var.get():
            raise ValueError("Status is required")
        if self.change_type_var.get() in (ChangeType.CONVERSION.value, ChangeType.TERMINATION.value) and not self._pc_employee_id:
            raise ValueError("Please select an employee for termination or conversion")
        try:
            effective_date = datetime.strptime(self.effective_date_var.get(), "%m/%d/%y").date()
        except ValueError:
            raise ValueError("Invalid Effective Date format. Use MM/DD/YY.")
        return {
            "description": self.description_var.get().strip(),
            "change_type": self.change_type_var.get(),
            "effective_date": effective_date,
            "name": self.name_var.get().strip(),
            "team": self.team_var.get().strip(),
            "manager_code": self.manager_code_var.get().strip(),
            "cost_center": self.cost_center_var.get().strip(),
            "employment_type": self.employment_type_var.get(),
            "status": self.status_var.get(),
            "employee_id": self._pc_employee_id,
        }
    
    def _save_detail(self):
        try:
            data = self._collect_detail_payload()
        except ValueError as e:
            messagebox.showerror("Error", str(e))
            return
        
        try:
            session = get_session()
            if self._record_id is None:
                change = PlannedChange(
                    description=data["description"],
                    change_type=data["change_type"],
                    effective_date=data["effective_date"],
                    name=data["name"],
                    team=data["team"],
                    manager_code=data["manager_code"],
                    cost_center=data["cost_center"],
                    employment_type=data["employment_type"],
                    status=data["status"],
                )
                if data.get("employee_id"):
                    change.employee_id = data["employee_id"]
                session.add(change)
                session.commit()
                new_id = change.id
                session.close()
                self.load_changes()
                self._select_tree_row_by_id(new_id)
                messagebox.showinfo("Success", "Planned change added successfully.")
            else:
                change = session.query(PlannedChange).filter(PlannedChange.id == self._record_id).first()
                if not change:
                    session.close()
                    messagebox.showerror("Error", "Planned change not found.")
                    return
                change.description = data["description"]
                change.change_type = data["change_type"]
                change.effective_date = data["effective_date"]
                change.name = data["name"]
                change.team = data["team"]
                change.manager_code = data["manager_code"]
                change.cost_center = data["cost_center"]
                change.employment_type = data["employment_type"]
                change.status = data["status"]
                if data.get("employee_id"):
                    change.employee_id = data["employee_id"]
                else:
                    change.employee_id = None
                session.commit()
                session.close()
                self.load_changes()
                self._select_tree_row_by_id(self._record_id)
                messagebox.showinfo("Success", "Planned change updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save planned change: {str(e)}")
    
    def _select_tree_row_by_id(self, change_id):
        self._suppress_tree_select = True
        try:
            for item in self.tree.get_children():
                if self.tree.item(item)["values"][0] == change_id:
                    self.tree.selection_set(item)
                    self.tree.see(item)
                    self._load_change_by_id(change_id)
                    return
        finally:
            self._suppress_tree_select = False
    
    def _cancel_detail(self):
        if self._record_id is not None:
            self._load_change_by_id(self._record_id)
        else:
            sel = self.tree.selection()
            if sel:
                self._on_tree_select()
            else:
                self._clear_detail_form(new_mode=True)
    
    def load_changes(self):
        """Load planned changes from database"""
        try:
            keep_id = self._record_id
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            session = get_session()
            year = int(self.year_var.get())
            changes = session.query(PlannedChange).filter(
                PlannedChange.effective_date.between(
                    datetime(year, 1, 1).date(),
                    datetime(year, 12, 31).date(),
                )
            ).all()
            
            count = 0
            for change in changes:
                item_id = self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        change.id,
                        change.description,
                        change.change_type,
                        change.effective_date.strftime("%m/%d/%y") if change.effective_date else "",
                        change.name or "",
                        change.team or "",
                        change.manager_code or "",
                        change.status,
                    ),
                )
                self.tree.item(item_id, tags=("evenrow",) if count % 2 == 1 else ("oddrow",))
                count += 1
            
            self.tree.tag_configure("oddrow", background=COLORS["white"])
            self.tree.tag_configure("evenrow", background=COLORS["table_row_alt"])
            session.close()
            
            if keep_id is not None:
                if any(self.tree.item(i)["values"][0] == keep_id for i in self.tree.get_children()):
                    self._select_tree_row_by_id(keep_id)
                else:
                    self._clear_detail_form(new_mode=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load planned changes: {str(e)}")
    
    def add_change(self):
        self._suppress_tree_select = True
        try:
            self.tree.selection_remove(self.tree.selection())
        except tk.TclError:
            pass
        self._suppress_tree_select = False
        self._clear_detail_form(new_mode=True)
    
    def edit_change(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a planned change to edit.")
            return
        change_id = self.tree.item(selected[0])["values"][0]
        self._load_change_by_id(change_id)
    
    def delete_change(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a planned change to delete.")
            return
        
        if not messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this planned change?"):
            return
        
        try:
            change_id = self.tree.item(selected[0])["values"][0]
            session = get_session()
            change = session.query(PlannedChange).filter(PlannedChange.id == change_id).first()
            if change:
                session.delete(change)
                session.commit()
            session.close()
            
            if self._record_id == change_id:
                self._clear_detail_form(new_mode=True)
            
            self.load_changes()
            messagebox.showinfo("Success", "Planned change deleted successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete planned change: {str(e)}")

class SettingsTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding="20")
        self.parent = parent
        
        # Create main frame with some padding
        main_frame = ttk.Frame(self, style='Card.TFrame', padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="Application Settings", 
                               font=("Helvetica", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Settings form
        settings_frame = ttk.Frame(main_frame)
        settings_frame.pack(fill=tk.X, pady=10)
        
        # FTE Hours
        fte_frame = ttk.Frame(settings_frame)
        fte_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(fte_frame, text="FTE Hours per Week:", width=20).pack(side=tk.LEFT, padx=5)
        self.fte_hours_var = tk.StringVar()
        fte_entry = ttk.Entry(fte_frame, textvariable=self.fte_hours_var, width=10)
        fte_entry.pack(side=tk.LEFT, padx=5)
        
        # Contractor Hours
        contractor_frame = ttk.Frame(settings_frame)
        contractor_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(contractor_frame, text="Contractor Hours per Week:", width=20).pack(side=tk.LEFT, padx=5)
        self.contractor_hours_var = tk.StringVar()
        contractor_entry = ttk.Entry(contractor_frame, textvariable=self.contractor_hours_var, width=10)
        contractor_entry.pack(side=tk.LEFT, padx=5)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Save Settings", command=self.save_settings, 
                  style="Primary.TButton", width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_defaults, 
                  width=15).pack(side=tk.LEFT, padx=5)
        
        # Load current settings
        self.load_settings()
    
    def load_settings(self):
        """Load settings from database"""
        try:
            session = get_session()
            settings = session.query(Settings).first()
            
            if not settings:
                # Create default settings if none exist
                settings = Settings(fte_hours=34.5, contractor_hours=39.0)
                session.add(settings)
                session.commit()
            
            # Set values in form
            self.fte_hours_var.set(str(settings.fte_hours))
            self.contractor_hours_var.set(str(settings.contractor_hours))
            
            session.close()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load settings: {str(e)}")
    
    def save_settings(self):
        """Save settings to database"""
        try:
            # Validate inputs
            try:
                fte_hours = float(self.fte_hours_var.get())
                contractor_hours = float(self.contractor_hours_var.get())
                
                if fte_hours <= 0 or contractor_hours <= 0:
                    raise ValueError("Hours must be positive numbers")
                    
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers for hours")
                return
            
            # Save to database
            session = get_session()
            settings = session.query(Settings).first()
            
            if not settings:
                settings = Settings()
                session.add(settings)
            
            settings.fte_hours = fte_hours
            settings.contractor_hours = contractor_hours
            
            session.commit()
            session.close()
            
            messagebox.showinfo("Success", "Settings saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")
    
    def reset_defaults(self):
        """Reset settings to default values"""
        self.fte_hours_var.set("34.5")
        self.contractor_hours_var.set("39.0")

class ForecastApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TSC Headcount")
        self.geometry("1200x800")
        self.configure(background=COLORS['background'])
        
        # Configure styles
        configure_styles()
        
        # Create a main frame to hold everything
        main_frame = ttk.Frame(self, style='Main.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.employee_tab = EmployeeTab(self.notebook)
        self.notebook.add(self.employee_tab, text="Employees")
        
        self.allocation_tab = ProjectAllocationTab(self.notebook)
        self.notebook.add(self.allocation_tab, text="Allocations")
        
        self.forecast_tab = ForecastTab(self.notebook)
        self.notebook.add(self.forecast_tab, text="Forecast")
        
        self.planned_changes_tab = PlannedChangesTab(self.notebook) 
        self.notebook.add(self.planned_changes_tab, text="Planned Changes")
        
        self.visualization_tab = ForecastVisualization(self.notebook)
        self.notebook.add(self.visualization_tab, text="Visualization")
        
        self.settings_tab = SettingsTab(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings")
        
        # Create status bar
        status_frame = ttk.Frame(self, style='Card.TFrame')
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=10, pady=5)
        
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT, padx=5)
        
        # Check database connection
        if not verify_db_connection():
            messagebox.showerror("Database Error", "Failed to connect to database.")
            self.status_var.set("Database connection failed")
        else:
            self.status_var.set("Connected to database")


if __name__ == "__main__":
    # Create database tables if they don't exist
    Base.metadata.create_all(engine)
    
    # Create the application
    app = ForecastApp()
    app.mainloop()
