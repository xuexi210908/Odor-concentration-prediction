
"""
This program predicts the odor concentration near the nose when the user is stationary,
at a position relative to the odor source (0, 0, 3) (units: mm).
The existing normalized alcohol PD data at 47°C can be imported into the UI for calculation.
"""
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
from openpyxl import load_workbook
from scipy.signal import butter, lfilter, lfilter_zi
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

PD_XLSX_PATH = Path(__file__).with_name("PD.xlsx")
FS = 10.0
CUTOFF_HZ = 0.01
TAU_C = 4.0
NOSE_X, NOSE_Y, NOSE_Z = 0.0, 0.0, 3.0
INITIAL_X_SECONDS = 5.0  # Initially display 0-5 s; then auto-expand with data

odor_params = {
    1: {  # Ethanol
        47: {"a": 6.1, "b": -15.0, "c": 8.6, "U_outlet": 0.0130, "M_molar": 46.07,"Q":0.03117},
        42: {"a": 5.7, "b": -11.8, "c": 6.4, "U_outlet": 0.0099, "M_molar": 46.07},
        37: {"a": 6.2, "b": -17.3, "c": 11.6, "U_outlet": 0.0075, "M_molar": 46.07},
        32: {"a": 5.8, "b": -16.0, "c": 11.3, "U_outlet": 0.0057, "M_molar": 46.07},
    },
    2: {  # Litchi
        47: {"a": 3.5, "b": -16.3, "c": 12.6, "U_outlet": 0.0070, "M_molar": 154.25,"Q":0.01537},
        42: {"a": 3.5, "b": -17.4, "c": 15.1, "U_outlet": 0.0052, "M_molar": 154.25},
        37: {"a": 3.2, "b": -18.5, "c": 18.1, "U_outlet": 0.0039, "M_molar": 154.25},
        32: {"a": 3.2, "b": -16.9, "c": 17.2, "U_outlet": 0.0029, "M_molar": 154.25},
    },
    3: {  # Pear
        47: {"a": 3.1, "b": -15.4, "c": 19.3, "U_outlet": 0.0080, "M_molar": 144.21,"Q":0.01514},
        42: {"a": 3.5, "b": -12.5, "c": 3.5, "U_outlet": 0.0059, "M_molar": 144.21},
        37: {"a": 3.4, "b": -11.8, "c": 2.7, "U_outlet": 0.0044, "M_molar": 144.21},
        32: {"a": 3.3, "b": -12.4, "c": 1.25, "U_outlet": 0.0032, "M_molar": 144.21},
    },
    4: {  # Carnation
        47: {"a": 2.5, "b": -10.5, "c": 10.4, "U_outlet": 0.0073, "M_molar": 164.2,"Q":0.01665},
        42: {"a": 2.9, "b": -6.0, "c": 0.5, "U_outlet": 0.0054, "M_molar": 164.2},
        37: {"a": 3.1, "b": -11.3, "c": -0.4, "U_outlet": 0.0040, "M_molar": 164.2},
        32: {"a": 3.2, "b": -8.8, "c": -10.3, "U_outlet": 0.0030, "M_molar": 164.2},
    },
    5: {  # Grass
        47: {"a": 2.8, "b": -9.1, "c": 3.0, "U_outlet": 0.0083, "M_molar": 98.15,"Q":0.01558},
        42: {"a": 3.0, "b": -8.3, "c": -3.0, "U_outlet": 0.0061, "M_molar": 98.15},
        37: {"a": 2.8, "b": -3.0, "c": 30.4, "U_outlet": 0.0045, "M_molar": 98.15},
        32: {"a": 3.0, "b": -10.5, "c": -6.3, "U_outlet": 0.0033, "M_molar": 98.15},
    },
    6: {  # Sandalwood
        47: {"a": 2.7, "b": -8.2, "c": 1.9, "U_outlet": 0.0075, "M_molar": 220.35,"Q":0.01566},
        42: {"a": 2.5, "b": -13.5, "c": 15.6, "U_outlet": 0.0055, "M_molar": 220.35},
        37: {"a": 2.5, "b": -11.2, "c": 13.2, "U_outlet": 0.0041, "M_molar": 220.35},
        32: {"a": 2.8, "b": -16.6, "c": 20.2, "U_outlet": 0.0030, "M_molar": 220.35},
    },
    7: {  # Red bean
        47: {"a": 4.3, "b": -12.2, "c": -2.6, "U_outlet": 0.0132, "M_molar": 100.16,"Q":0.03109},
        42: {"a": 4.4, "b": -17.3, "c": 9.7, "U_outlet": 0.0097, "M_molar": 100.16},
        37: {"a": 4.3, "b": -16.0, "c": 8.0, "U_outlet": 0.0072, "M_molar": 100.16},
        32: {"a": 4.4, "b": -13.6, "c": 3.4, "U_outlet": 0.0054, "M_molar": 100.16},
    },
    8: {  # Fir bark
        47: {"a": 3.6, "b": -13.5, "c": 6.2, "U_outlet": 0.0089, "M_molar": 136.24,"Q":0.01808},
        42: {"a": 3.7, "b": -17.1, "c": 9.5, "U_outlet": 0.0066, "M_molar": 136.24},
        37: {"a": 3.7, "b": -20.3, "c": 22.7, "U_outlet": 0.0048, "M_molar": 136.24},
        32: {"a": 3.5, "b": -8.6, "c": -11.7, "U_outlet": 0.0036, "M_molar": 136.24},
    },
    9: {  # Loquat
        47: {"a": 3.7, "b": -3.5, "c": -28.6, "U_outlet": 0.0087, "M_molar": 100.16,"Q":0.02317},
        42: {"a": 4.1, "b": -5.5, "c": -34.2, "U_outlet": 0.0064, "M_molar": 100.16},
        37: {"a": 3.9, "b": -6.3, "c": -17.3, "U_outlet": 0.0047, "M_molar": 100.16},
        32: {"a": 4.1, "b": -7.2, "c": -3.0, "U_outlet": 0.0035, "M_molar": 100.16},
    },
    10: {  # Chocolate
        47: {"a": 2.8, "b": -3.3, "c": -18.9, "U_outlet": 0.0076, "M_molar": 122.17,"Q":0.01673},
        42: {"a": 3.1, "b": -4.3, "c": -15.2, "U_outlet": 0.0056, "M_molar": 122.17},
        37: {"a": 3.3, "b": -5.4, "c": -12.3, "U_outlet": 0.0042, "M_molar": 122.17},
        32: {"a": 3.4, "b": -19.3, "c": 28.3, "U_outlet": 0.0031, "M_molar": 122.17},
    },
    11: {  # Orange
        47: {"a": 3.5, "b": -8.3, "c": -13.4, "U_outlet": 0.0088, "M_molar": 136.24,"Q":0.02056},
        42: {"a": 3.8, "b": -4.6, "c": -18.5, "U_outlet": 0.0065, "M_molar": 136.24},
        37: {"a": 4.1, "b": -4.7, "c": -26.4, "U_outlet": 0.0048, "M_molar": 136.24},
        32: {"a": 3.7, "b": -7.8, "c": -11.0, "U_outlet": 0.0035, "M_molar": 136.24},
    },
    12: {  # Passionfruit
        47: {"a": 4.0, "b": -6.4, "c": -57.9, "U_outlet": 0.0091, "M_molar": 150.22,"Q":0.02175},
        42: {"a": 3.5, "b": -17.3, "c": 5.4, "U_outlet": 0.0067, "M_molar": 150.22},
        37: {"a": 4.2, "b": -8.9, "c": -2.4, "U_outlet": 0.0050, "M_molar": 150.22},
        32: {"a": 4.1, "b": -14.3, "c": 5.1, "U_outlet": 0.0037, "M_molar": 150.22},
    },
    13: {  # Ice cream
        47: {"a": 4.2, "b": -12.9, "c": -2.5, "U_outlet": 0.0094, "M_molar": 170.25,"Q":0.02202},
        42: {"a": 4.6, "b": -13.3, "c": -0.73, "U_outlet": 0.0069, "M_molar": 170.25},
        37: {"a": 4.7, "b": -14.8, "c": 6.2, "U_outlet": 0.0051, "M_molar": 170.25},
        32: {"a": 4.3, "b": -13.4, "c": 2.4, "U_outlet": 0.0037, "M_molar": 170.25},
    },
    14: {  # Bamboo
        47: {"a": 4.4, "b": -9.0, "c": -6.3, "U_outlet": 0.0082, "M_molar": 98.15,"Q":0.01999},
        42: {"a": 4.4, "b": -8.6, "c": -8.2, "U_outlet": 0.0060, "M_molar": 98.15},
        37: {"a": 4.1, "b": -13.4, "c": -1.9, "U_outlet": 0.0045, "M_molar": 98.15},
        32: {"a": 4.3, "b": -18.9, "c": 9.9, "U_outlet": 0.0033, "M_molar": 98.15},
    },
    15: {  # White orchid
        47: {"a": 3.6, "b": -5.4, "c": -12.5, "U_outlet": 0.0085, "M_molar": 122.17,"Q":0.01664},
        42: {"a": 3.5, "b": -5.8, "c": -11.6, "U_outlet": 0.0063, "M_molar": 122.17},
        37: {"a": 3.7, "b": -8.7, "c": -3.9, "U_outlet": 0.0047, "M_molar": 122.17},
        32: {"a": 3.6, "b": -8.1, "c": -5.2, "U_outlet": 0.0034, "M_molar": 122.17},
    },
    16: {  # Beer
        47: {"a": 3.5, "b": -9.5, "c": -9.1, "U_outlet": 0.0079, "M_molar": 130.19,"Q":0.01759},
        42: {"a": 3.1, "b": -11.8, "c": 8.1, "U_outlet": 0.0058, "M_molar": 130.19},
        37: {"a": 3.5, "b": -14.5, "c": 11.6, "U_outlet": 0.0043, "M_molar": 130.19},
        32: {"a": 3.4, "b": -20.5, "c": 22.7, "U_outlet": 0.0032, "M_molar": 130.19},
    },

}


def gaussian_plume(U, x, y, z, k, sigma_y, sigma_z, h, Q=0.025, H=0, M=46.07):
    """
    calculate the odor concentration based on the Gaussian plume model.
    U: Composite wind speed magnitude (m/s)
    x: Downwind coordinate (m)
    y: Crosswind coordinate (m)
    z: Vertical coordinate (m)
    k, h: Correction coefficients
    sigma_y: Crosswind diffusion coefficient (m)
    sigma_z: Vertical diffusion coefficient (m)
    Q: Odor emission rate (mg/s)
    H: Odor source height (m)
    M: Molar mass of the odor (g/mol)
    :return: Odor concentration at the point (ppm)
    """
    term1 = Q / (2 * np.pi * U * sigma_y * sigma_z)
    term2 = np.exp(-y ** 2 / (2 * sigma_y ** 2))

    term3 = np.exp(-((z - H) ** 2) / (2 * sigma_z ** 2))  # Ground reflection is not considered
    return k * (
                term1 * term2 * term3 * 24.5 / M) + h  # Convert the odor concentration unit to ppm, where 24.5 is the molar volume of air at 25°C and 1 atm


def transform_coordinates_two_steps(x, y, z, theta1, theta2):
    """
    Transform the coordinates of the calculation point from the original coordinate system to the coordinates under the Gaussian plume model.
    Two sequential rotations of the coordinate axes:
    First rotation: around the Z-axis (rotation in the XY plane by theta1)
    Second rotation: around the Y-axis (rotation in the XZ plane by theta2)
    Parameters:
    x, y, z: Original coordinates
    theta1: Rotation angle around the Z-axis (unit: radians)
    theta2: Rotation angle around the Y-axis (unit: radians)
    Returns:
    x_new, y_new, z_new: Transformed new coordinates
    """

    # First rotation: around the Z-axis
    x1 = x * np.cos(theta1) + y * np.sin(theta1)
    y1 = -x * np.sin(theta1) + y * np.cos(theta1)
    z1 = z

    # Second rotation: around the Y-axis
    x_new = x1 * np.cos(theta2) + z1 * np.sin(theta2)
    y_new = y1
    z_new = -x1 * np.sin(theta2) + z1 * np.cos(theta2)

    return x_new, y_new, z_new


# .......................................................................................................
def gaussian_model(Walk_x, Walk_y, U_outlet, nose_x, nose_y, nose_z, Rate_OdorRelease, M_Odor, a, b, sigma0, k, h):
    """
    Main
    :param Walk_x,Walk_y: Walking speed of the human body on the horizontal plane (m/s)
    :param nose_x,nose_y,nose_z: Coordinates of the human nose (mm)
    :param Rate_OdorRelease: Odor release rate of the odor generator (OG) (mg/s)
    :param M_Odor: Molar mass of the odor (g/mol)
    :return:
    """
    U_x = -Walk_x
    U_y = -Walk_y
    U_z = U_outlet
    x, y, z = nose_x * 0.001, nose_y * 0.001, nose_z * 0.001
    # Calculate the resultant speed U_xy and direction theta1 (radians) on the horizontal plane
    U_xy = np.sqrt(U_x ** 2 + U_y ** 2)
    theta1 = np.arctan2(U_y, U_x)
    # Calculate the resultant speed U and direction theta2 (radians)
    U = np.sqrt(U_xy ** 2 + U_z ** 2)
    theta2 = np.arctan2(U_z, U_xy)
    x_rot, y_rot, z_rot = transform_coordinates_two_steps(x, y, z, theta1, theta2)
    sigma_yz = a * (np.abs(x_rot) ** b) + sigma0
    Concen = gaussian_plume(U, x_rot, y_rot, z_rot, k, sigma_yz, sigma_yz, h, Q=Rate_OdorRelease, H=0, M=M_Odor)
    return Concen

ODOR_NAMES = {
    1: "Ethanol", 2: "Litchi", 3: "Pear", 4: "Carnation",
    5: "Grass", 6: "Sandalwood", 7: "Red bean", 8: "Fir bark",
    9: "Loquat", 10: "Chocolate", 11: "Orange", 12: "Passionfruit",
    13: "Ice cream", 14: "Bamboo", 15: "White orchid", 16: "Beer"
}


def read_pd_xlsx(path):
    """Read the t and pd columns from the Excel file; uses the first worksheet by default. PD is already normalized."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        iterator = sheet.iter_rows(values_only=True)
        first_row = next(iterator, None)
        if first_row is None:
            raise ValueError("Excel worksheet is empty")

        # Prefer identifying t and pd by header; otherwise default to the first two columns.
        names = [str(v).strip().lower() if v is not None else "" for v in first_row]
        if "t" in names and "pd" in names:
            t_col, pd_col = names.index("t"), names.index("pd")
            rows = iterator
            start_row = 2
        else:
            t_col, pd_col = 0, 1
            from itertools import chain
            rows = chain((first_row,), iterator)
            start_row = 1

        data = []
        for row_number, row in enumerate(rows, start_row):
            if len(row) <= max(t_col, pd_col):
                continue
            t_value, pd_value = row[t_col], row[pd_col]
            if t_value is None and pd_value is None:
                continue
            try:
                ti, pd = float(t_value), float(pd_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Excel row {row_number}: t or pd is not a numeric value") from exc
            if not (np.isfinite(ti) and np.isfinite(pd)):
                raise ValueError(f"Excel row {row_number} contains non-finite values")
            data.append((ti, pd))
    finally:
        workbook.close()

    if len(data) < 2:
        raise ValueError("Excel requires at least two valid samples")
    data = np.asarray(data, dtype=float)
    if np.any(np.diff(data[:, 0]) <= 0):
        raise ValueError("Excel time t must be strictly increasing")
    return data[:, 0], data[:, 1]



class Channel:
    def __init__(self, app, number, parent, color):
        self.app = app
        self.number = number
        self.path = PD_XLSX_PATH
        self.running = False
        self.after_id = None
        self.times = np.array([])
        self.pds = np.array([])
        self.index = 0
        self.plot_t, self.plot_c = [], []
        self.previous_concentration = 40.0
        self.last_mass = None
        self.filter_state = None
        self.last_time = None
        self.bf, self.af = butter(2, CUTOFF_HZ / (FS / 2), btype="low")
        self.line, = app.ax.plot([], [], color=color, linewidth=1.5, label=f"CH{number:02d}")
        self.status = tk.StringVar(value="Not running")
        self.concentration = tk.StringVar(value="--")
        self.rate = tk.StringVar(value="--")
        self.file_name = tk.StringVar(value="PD.xlsx")
        ttk.Label(parent, text=f"{number:02d}", width=3).grid(row=number, column=0, padx=3, pady=3)
        self.odor = ttk.Combobox(parent, state="readonly", width=17,
            values=[f"{i:02d} - {ODOR_NAMES[i]}" for i in range(1,17)])
        self.odor.current(number-1)
        self.odor.grid(row=number, column=1, padx=2)
        self.temp = ttk.Combobox(parent, state="readonly", width=5, values=["32", "37", "42", "47"])
        self.temp.set("47")
        self.temp.grid(row=number, column=2, padx=2)
        ttk.Button(parent, text="Select", width=5, command=self.choose_file).grid(row=number, column=3, padx=2)
        ttk.Label(parent, textvariable=self.file_name, width=13).grid(row=number, column=4, padx=2)
        self.start_btn = ttk.Button(parent, text="Run", width=5, command=self.start)
        self.start_btn.grid(row=number, column=5, padx=2)
        ttk.Button(parent, text="Stop", width=5, command=self.stop).grid(row=number, column=6, padx=2)
        ttk.Label(parent, textvariable=self.concentration, width=12).grid(row=number, column=7, padx=2)
        ttk.Label(parent, textvariable=self.rate, width=13).grid(row=number, column=8, padx=2)
        ttk.Label(parent, textvariable=self.status, width=13).grid(row=number, column=9, padx=2)

    def choose_file(self):
        selected = filedialog.askopenfilename(title=f"Channel {self.number:02d}: Select PD Excel",
                                               filetypes=[("Excel files", "*.xlsx")])
        if selected:
            self.path = Path(selected)
            self.file_name.set(self.path.name)

    def start(self):
        self.stop(update_status=False)
        try:
            if not self.path.is_file():
                self.choose_file()
                if not self.path.is_file():
                    self.status.set("Missing Excel")
                    return
            self.times, self.pds = read_pd_xlsx(self.path)
            odor_id = int(self.odor.get().split(" - ")[0])
            self.params = odor_params[odor_id][int(self.temp.get())]
        except Exception as exc:
            self.status.set("Start failed")
            messagebox.showerror(f"Channel {self.number:02d}", str(exc))
            return
        self.index = 0
        self.plot_t, self.plot_c = [], []
        self.previous_concentration = 40.0
        self.last_mass = None
        self.filter_state = None
        self.last_time = None
        self.line.set_data([], [])
        self.concentration.set("--")
        self.rate.set("--")
        self.running = True
        self.odor.configure(state="disabled")
        self.temp.configure(state="disabled")
        self.start_btn.configure(state="disabled")
        self.status.set("Running")
        self.app.update_plot()
        self.tick()

    def tick(self):
        self.after_id = None
        if not self.running:
            return
        try:
            i = self.index
            t = float(self.times[i])
            pd = float(self.pds[i])
            p = self.params
            mass = p["a"] + p["b"] * pd + p["c"] * pd ** 2
            if self.filter_state is None:
                self.filter_state = lfilter_zi(self.bf, self.af) * mass
                self.last_mass = mass
            filtered, self.filter_state = lfilter(self.bf, self.af, [mass], zi=self.filter_state)
            current_mass = float(filtered[0])
            dt = 1.0 / FS if self.last_time is None else t - self.last_time
            q = max(0.0, (current_mass - self.last_mass) / dt)
            self.last_mass = current_mass
            self.last_time = t
            c0 = gaussian_model(0, 0, p["U_outlet"], NOSE_X, NOSE_Y, NOSE_Z,
                                q, p["M_molar"], 1.64, 1.18, 0.0023, 0.14, 40.0)
            target = 0.4618 * c0 + 60 if c0 > 92.15 else 1.1129 * c0
            self.previous_concentration += dt / TAU_C * (target - self.previous_concentration)
            c = float(self.previous_concentration)
            self.plot_t.append(t - float(self.times[0]))
            self.plot_c.append(c)
            self.line.set_data(self.plot_t, self.plot_c)
            self.concentration.set(f"{c:.4f}")
            self.rate.set(f"{q:.6f}")
            self.index += 1
            self.app.request_plot_update()
            if self.index >= len(self.times):
                self.stop(update_status=False)
                self.status.set("Completed")
                return
            next_dt = float(self.times[self.index] - self.times[i])
            self.after_id = self.app.root.after(max(1, round(next_dt * 1000)), self.tick)
        except Exception as exc:
            self.stop(update_status=False)
            self.status.set("Calculation failed")
            messagebox.showerror(f"Channel {self.number:02d} calculation error", str(exc))

    def stop(self, update_status=True):
        self.running = False
        if self.after_id is not None:
            self.app.root.after_cancel(self.after_id)
            self.after_id = None
        self.odor.configure(state="readonly")
        self.temp.configure(state="readonly")
        self.start_btn.configure(state="normal")
        if update_status:
            self.status.set("Stopped")

class OG16App:
    def __init__(self, root):
        self.root = root
        root.title("16-Channel OG Odor Concentration Prediction")
        root.geometry("1500x900")
        root.minsize(1100, 650)
        self.plot_after_id = None
        self.channels = []
        bar = ttk.Frame(root, padding=8)
        bar.pack(fill="x")
        ttk.Label(bar, text="16-Channel OG Odor Concentration Prediction", font=("Arial", 15, "bold")).pack(side="left", padx=8)
        ttk.Button(bar, text="Run All", command=self.start_all).pack(side="left", padx=8)
        ttk.Button(bar, text="Stop All", command=self.stop_all).pack(side="left", padx=8)
        container = ttk.PanedWindow(root, orient="vertical")
        container.pack(fill="both", expand=True, padx=10, pady=6)
        table_frame = ttk.Frame(container)
        container.add(table_frame, weight=2)
        scroll_canvas = tk.Canvas(table_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=scroll_canvas.yview)
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        scroll_canvas.pack(side="left", fill="both", expand=True)
        rows = ttk.Frame(scroll_canvas)
        win = scroll_canvas.create_window((0,0), window=rows, anchor="nw")
        rows.bind("<Configure>", lambda e: scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all")))
        scroll_canvas.bind("<Configure>", lambda e: scroll_canvas.itemconfigure(win, width=e.width))
        for j, name in enumerate(["Channel", "Odor Type", "Temp °C", "Excel", "File Name", "Run", "Stop", "Concentration (ppm)", "Q (mg/s)", "Status"]):
            ttk.Label(rows, text=name, font=("Arial", 10, "bold")).grid(row=0,column=j,padx=3,pady=7)
        plot_frame = ttk.Frame(container)
        container.add(plot_frame, weight=3)
        self.figure = Figure(figsize=(12,5), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_title("Predicted Odor Concentration - 16 Channels")
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Concentration (ppm)")
        self.ax.grid(True, alpha=0.3)
        self.ax.set_xlim(0, INITIAL_X_SECONDS)
        self.ax.set_ylim(0, 100)
        colors = __import__('matplotlib').colormaps['tab20'].colors
        for i in range(16):
            self.channels.append(Channel(self, i+1, rows, colors[i]))
        self.ax.legend(loc="upper left", ncol=8, fontsize=8)
        self.figure.tight_layout()
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def request_plot_update(self):
        if self.plot_after_id is None:
            self.plot_after_id = self.root.after(100, self.update_plot)

    def update_plot(self):
        self.plot_after_id = None
        last_times = [ch.plot_t[-1] for ch in self.channels if ch.plot_t]
        if last_times:
            tmax = max(last_times)
            # Auto-expand: show all history from 0 s onward, no scrolling or clipping
            self.ax.set_xlim(0, max(INITIAL_X_SECONDS, tmax + 1.0))
            visible = [c for ch in self.channels for c in ch.plot_c]
            if visible:
                low, high = min(visible), max(visible)
                pad = max(5.0, (high-low)*0.1)
                self.ax.set_ylim(max(0, low-pad), high+pad)
        else:
            self.ax.set_xlim(0, INITIAL_X_SECONDS)
            self.ax.set_ylim(0, 100)
        self.canvas.draw_idle()

    def start_all(self):
        for ch in self.channels:
            ch.start()

    def stop_all(self):
        for ch in self.channels:
            ch.stop()
        self.update_plot()

    def on_exit(self):
        for ch in self.channels:
            ch.stop(update_status=False)
        if self.plot_after_id is not None:
            self.root.after_cancel(self.plot_after_id)
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    OG16App(root)
    root.mainloop()