import pandas as pd
import numpy as np
from scipy.optimize import minimize
import tkinter as tk
from tkinter import filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
from tkinter import ttk
import os

# LabelTable class for displaying a table
class LabelTable(ttk.Frame):
    def __init__(self, parent, rows=2, columns=2, cell_width=10, cell_height=1, cell_font=("Helvetica", 10), data=None):
        ttk.Frame.__init__(self, parent)
        self._widgets = []
        self.rows = rows
        self.columns = columns
        self.cell_width = cell_width
        self.cell_height = cell_height
        self.cell_font = cell_font

        for row in range(self.rows):
            current_row = []
            for column in range(self.columns):
                label = tk.Label(self, text='', width=self.cell_width, height=self.cell_height, font=self.cell_font,
                                 borderwidth=1, relief='solid')
                label.grid(row=row, column=column, sticky="nsew", padx=1, pady=1)
                current_row.append(label)
            self._widgets.append(current_row)

        if data:
            self.set_data(data)

        for row in range(self.rows):
            self.grid_rowconfigure(row, weight=1)

        for col in range(self.columns):
            self.grid_columnconfigure(col, weight=1)

    def set_data(self, data):
        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                formatted_value = f"{value:.6f}" if isinstance(value, (float, np.floating)) else str(value)
                self._widgets[i][j].config(text=formatted_value)

# Define λ_fixed as a global variable
λ_fixed = None

# Function to run the optimization and generate output
def run_optimization(file_path, β0, β1, β2, λ, plot_window, table_frame):
    global λ_fixed  # Use the global variable

    # Load data
    dd = pd.read_excel(file_path)

    def myval(c):
        dd = pd.read_excel(file_path)
        c[3] = λ_fixed if λ_fixed is not None else c[3]
        ddd = dd.copy()
        ddd['NS'] = (c[0]) + (c[1] * ((1 - np.exp(-ddd.iloc[:, 0] / c[3])) / (ddd.iloc[:, 0] / c[3]))) + (
                c[2] * ((((1 - np.exp(-ddd.iloc[:, 0] / c[3])) / (ddd.iloc[:, 0] / c[3]))) - (np.exp(-ddd.iloc[:, 0] / c[3]))))
        ddd['Residual'] = (ddd.iloc[:, 1] - ddd.iloc[:, 2]) ** 2
        val = np.sum(ddd.iloc[:, 3])
        print("[β0, β1, β2, λ]=", c, ", SUM:", val)
        return val

    bnds = ((0, None), (None, None), (None, None), (0.031, 0.032))  # λ within [0, 10]

    X0 = [β0, β1, β2, λ]
    c = minimize(myval, X0, method='SLSQP', bounds=bnds, options={'disp': True})

    β0 = c.x[0]
    β1 = c.x[1]
    β2 = c.x[2]
    λ = c.x[3]

    NS = pd.DataFrame(columns=['Maturity', 'Zero Coupon yield curve (cont.comp.)'])
    NS['Maturity'] = np.arange(0, 15, 1/365)

    # Compute Zero Coupon Yield Curve
    NS['Zero Coupon yield curve (cont.comp.)'] = (β0) + (β1 * ((1 - np.exp(-NS['Maturity'] / λ)) / (NS['Maturity'] / λ))) + (
            β2 * ((((1 - np.exp(-NS['Maturity'] / λ)) / (NS['Maturity'] / λ))) - (np.exp(-NS['Maturity'] / λ))))
    
    # Compute Implied Forward Curve
    NS['Implied Forward Curve'] = (((1+NS.iloc[:,1]/365)**(NS.iloc[:,0]*365))/((1+NS.iloc[:,1].shift(1)/365)**(NS.iloc[:,0].shift(1)*365))-1)*365
    NS['Implied Forward Curve'][0] = NS['Zero Coupon yield curve (cont.comp.)'][0]
    # Compute Discount Factor
    NS['Discount Factor'] = 1 / (1 + NS['Zero Coupon yield curve (cont.comp.)'] / 365) ** (365 * NS['Maturity'])

    # Compute Par Curve
    NS['Par Curve'] = (np.exp((1-NS.iloc[:,3])/np.cumsum(NS.iloc[:,3])*365*0.5)-1)/0.5
    NS['Par Curve'].iloc[:30] = NS['Zero Coupon yield curve (cont.comp.)'].iloc[:30]

    NS['Maturity (in days)'] = NS['Maturity']*365
    NS.to_excel('Implied NS curves (cont).xlsx', index=False)

    # Display optimization result in a dialog box
    if c.success:
        messagebox.showinfo("Optimization Complete", "Optimization successfully completed.")
        plot_ns(NS, plot_window, β0, β1, β2, λ, table_frame)
        save_implied_curve_results(NS)
        save_parameters_results(β0, β1, β2, λ)
        merge_and_save_sheets()
    else:
        messagebox.showerror("Optimization Error", "Optimization did not converge. Please check the parameters.")

# Function to handle the file dialog
def browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            # Load data
            dd = pd.read_excel(file_path)
            messagebox.showinfo("Data Upload", "Data has been successfully uploaded.")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Data Upload Error", "The selected file is empty. Please choose another file.")
            return
        except Exception as e:
            messagebox.showerror("Data Upload Error", f"An error occurred while uploading data:\n{str(e)}")
            return

        # Create a new window for parameter sliders
        param_window = tk.Toplevel(root)
        param_window.title("Parameter Controls")

        # Create sliders for parameters with larger dimensions
        β0_slider = tk.Scale(param_window, label="β0 (to be optimized by NS model)", from_=0, to=1, resolution=0.01,
                             orient='horizontal', length=400)
        β0_slider.set(0.01)
        β0_slider.pack(pady=10)

        β1_slider = tk.Scale(param_window, label="β1 (to be optimized by NS model)", from_=0, to=1, resolution=0.01,
                             orient='horizontal', length=400)
        β1_slider.set(0.01)
        β1_slider.pack(pady=10)

        β2_slider = tk.Scale(param_window, label="β2 (to be optimized by NS model)", from_=0, to=1, resolution=0.01,
                             orient='horizontal', length=400)
        β2_slider.set(0.01)
        β2_slider.pack(pady=10)

        λ_slider = tk.Scale(param_window, label="λ (fixed, within [0.031, 0.032])", from_=0, to=10, resolution=0.01,
                            orient='horizontal', length=400)
        λ_slider.set(0.03)
        λ_slider.pack(pady=10)

        # Button to run optimization
        run_button = tk.Button(param_window, text="Run Optimization",
                               command=lambda: run_optimization(file_path,
                                                                β0_slider.get(),
                                                                β1_slider.get(),
                                                                β2_slider.get(),
                                                                λ_slider.get(),
                                                                param_window,
                                                                table_frame))
        run_button.pack(pady=20)

# Function to plot the NS dataframe
def plot_ns(NS, plot_window, β0, β1, β2, λ, table_frame):
    # Create a new window for the plot and table
    plot_window.title("Yield Curve Plot - Optimized Parameters")

    # Create a matplotlib figure
    fig = Figure(figsize=(6, 5), dpi=100)
    ax = fig.add_subplot(1, 1, 1)

    ax.plot(NS['Maturity'], NS['Zero Coupon yield curve (cont.comp.)'], label='Zero Coupon Yield Curve')
    ax.plot(NS['Maturity'], NS['Implied Forward Curve'], label='Implied Forward Curve')
    ax.plot(NS['Maturity'], NS['Par Curve'], label='Par Curve', linestyle='--')

    ax.set_xlabel('Maturity')
    ax.set_ylabel('Yield')
    ax.legend()

    # Embed the matplotlib figure in the Tkinter window
    canvas = FigureCanvasTkAgg(fig, master=plot_window)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Add a navigation toolbar to the plot window
    toolbar = NavigationToolbar2Tk(canvas, plot_window)
    toolbar.update()
    canvas.get_tk_widget().pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

    # Create a frame for the table
    table_frame.destroy()  # Destroy previous frame if it exists
    table_frame = tk.Frame(plot_window)
    table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=1)

    # Create a table to display the estimated parameters
    table_data = [['Parameter', 'Estimated Value'],
                  ['β0', β0],
                  ['β1', β1],
                  ['β2', β2],
                  ['λ', λ]]
    table = LabelTable(table_frame, rows=len(table_data), columns=len(table_data[0]), data=table_data)
    table.pack(expand=tk.YES, fill=tk.BOTH)

# Function to save implied curve results
def save_implied_curve_results(NS):
    implied_curve_sheet_path = 'implied_curve_results.xlsx'
    
    # Check if the sheet already exists
    if not os.path.exists(implied_curve_sheet_path):
        # If it doesn't exist, create a new DataFrame with the required columns
        columns = ['Optimization Date', 'UZONIA', 'UZDC', '1 month', '3 months', '6 months', '9 months',
                   '1 year', '2 years', '3 years', '5 years', '10 years', '15 years']
        implied_curve_data = pd.DataFrame(columns=columns)
    else:
        # If it exists, load the existing DataFrame
        implied_curve_data = pd.read_excel(implied_curve_sheet_path)

    # Filter out and drop previous results from today
    today_date = pd.to_datetime('today').floor('D')
    implied_curve_data = implied_curve_data[implied_curve_data['Optimization Date'] != today_date]
    
    # Add a new row with the current date and corresponding values from NS DataFrame
    values_to_match = [1/365, 7/365, 30/365, 90/365, 180/365, 270/365, 365/365, 730/365, 1095/365, 1825/365, 3650/365, 5474/365]
    new_row = [today_date] + list(NS.loc[NS['Maturity'].apply(lambda x: any(np.isclose(x, v) for v in values_to_match)),
                                         'Zero Coupon yield curve (cont.comp.)'].values)

    implied_curve_data = implied_curve_data.append(pd.Series(new_row, index=implied_curve_data.columns), ignore_index=True)

    # Save the updated DataFrame to the Excel file
    implied_curve_data.to_excel(implied_curve_sheet_path, index=False)

# Function to save parameters results
def save_parameters_results(β0, β1, β2, λ):
    parameters_sheet_path = 'parameters_results.xlsx'
    
    # Check if the sheet already exists
    if not os.path.exists(parameters_sheet_path):
        # If it doesn't exist, create a new DataFrame with the required columns
        columns = ['Optimization Date', 'Beta0', 'Beta1', 'Beta2', 'Lambda']
        parameters_data = pd.DataFrame(columns=columns)
    else:
        # If it exists, load the existing DataFrame
        parameters_data = pd.read_excel(parameters_sheet_path)

    # Filter out and drop previous results from today
    today_date = pd.to_datetime('today').floor('D')
    parameters_data = parameters_data[parameters_data['Optimization Date'] != today_date]

    # Add a new row with the current date and corresponding parameter values
    new_row = [today_date, β0, β1, β2, λ]
    parameters_data = parameters_data.append(pd.Series(new_row, index=parameters_data.columns), ignore_index=True)

    # Save the updated DataFrame to the Excel file
    parameters_data.to_excel(parameters_sheet_path, index=False)

# Function to merge and save all sheets into a single Excel file
def merge_and_save_sheets():
    implied_curve_sheet_path = 'implied_curve_results.xlsx'
    parameters_sheet_path = 'parameters_results.xlsx'
    output_excel_path = 'merged_results.xlsx'
    
    # Load existing implied curve and parameters sheets, if they exist
    implied_curve_data = pd.read_excel(implied_curve_sheet_path) if os.path.exists(implied_curve_sheet_path) else pd.DataFrame()
    parameters_data = pd.read_excel(parameters_sheet_path) if os.path.exists(parameters_sheet_path) else pd.DataFrame()
    
    # Merge the sheets and save to a new Excel file
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        implied_curve_data.to_excel(writer, sheet_name='Implied Zero Coupon curve', index=False)
        parameters_data.to_excel(writer, sheet_name='Parameters', index=False)

# Create the main Tkinter window
root = tk.Tk()
root.title("Yield Curve Optimizer")

# Create a button to browse for a file
browse_button = tk.Button(root, text="Browse File", command=browse_file)
browse_button.pack(pady=20)

# Initialize table frame
table_frame = tk.Frame(root)

# Run the Tkinter main loop
root.mainloop()
