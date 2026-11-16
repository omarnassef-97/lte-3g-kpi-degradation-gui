import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from tkinter import *
from tkinter import messagebox, filedialog

kpis_list = []

def update_kpi_menu():
    global kpis_list
    try:
        df = pd.read_excel(input_file_path_entry.get(), nrows=1)  # Read only the header row
        kpis_list = df.columns[2:].tolist()  # Extract KPIs from the header
        menu = kpi_menu["menu"]
        menu.delete(0, "end")
        for kpi in kpis_list:
            menu.add_command(label=kpi, command=lambda value=kpi: selected_kpi_var.set(value))
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def choose_input_file():
    input_file_path = filedialog.askopenfilename(title="Select Input Excel File")
    input_file_path_entry.delete(0, END)
    input_file_path_entry.insert(0, input_file_path)
    update_kpi_menu()

def choose_output_file():
    output_file_path = filedialog.asksaveasfilename(title="Save As", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_file_path_entry.delete(0, END)
    output_file_path_entry.insert(0, output_file_path)

def generate_excel_3G_WBTS():
    input_file_path = input_file_path_entry.get()
    output_file_path = output_file_path_entry.get()
    selected_kpi = selected_kpi_var.get()
    start_date_str = entry_start_date.get()

    if not input_file_path or not output_file_path or not selected_kpi or not start_date_str:
        messagebox.showwarning("Warning", "Please select input file, output file, enter KPI, and select a start date.")
        return

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()

    df = pd.read_excel(input_file_path, index_col=False, sheet_name=0)

    time_period = int(time_period_entry.get())

    if time_period == 14:
        # Get data for the same day of the week for the last 2 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 3)])]

    elif time_period == 30:
        # Get data for the same day of the week for the last 4 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 5)])]

    unique_dates = df2["Period start time"].dt.date.unique()
    columns_count = len(unique_dates)

    compare = []
    comparison_operator = comparison_operator_var.get()

    for WBTS_name, group in df2.groupby("WBTS name"):
        orig_values = df.loc[
            (df["WBTS name"] == WBTS_name) & (df["Period start time"].dt.date == start_date), selected_kpi]

        for date in unique_dates:
            ref_value = group.loc[group["Period start time"].dt.date == date, selected_kpi]

            if not orig_values.empty and not ref_value.empty:
                if comparison_operator == "> 'EX: DR'":
                    is_degraded = all(orig_value > ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((WBTS_name, orig_values, is_degraded, date, difference))
                elif comparison_operator == "< 'EX:CSSR'":
                    is_degraded = all(orig_value < ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((WBTS_name, orig_values.values[0], is_degraded, date, difference))

    compare_df = pd.DataFrame(compare, columns=["WBTS name", "Current value","Degraded", "dates", "difference"])
    difference_range = compare_df.groupby("WBTS name")["difference"].agg(["min", "max"]).reset_index()
    compare_df = compare_df.merge(difference_range, on="WBTS name")

    compare_df_pivot = compare_df.pivot(index=["WBTS name", "Current value"], columns="dates",
                                        values=["Degraded", "difference"])

    compare_df_pivot.to_excel(output_file_path)
    messagebox.showinfo("Success", "Excel file generated successfully.")


def generate_excel_3G_WCEL():
    input_file_path = input_file_path_entry.get()
    output_file_path = output_file_path_entry.get()
    selected_kpi = selected_kpi_var.get()
    start_date_str = entry_start_date.get()

    if not input_file_path or not output_file_path or not selected_kpi or not start_date_str:
        messagebox.showwarning("Warning", "Please select input file, output file, enter KPI, and select a start date.")
        return

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()

    df = pd.read_excel(input_file_path, index_col=False, sheet_name=0)

    time_period = int(time_period_entry.get())

    if time_period == 14:
        # Get data for the same day of the week for the last 2 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 3)])]

    elif time_period == 30:
        # Get data for the same day of the week for the last 4 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 5)])]

    unique_dates = df2["Period start time"].dt.date.unique()
    columns_count = len(unique_dates)

    compare = []
    comparison_operator = comparison_operator_var.get()

    for WCEL_name, group in df2.groupby("WCEL name"):
        orig_values = df.loc[
            (df["WCEL name"] == WCEL_name) & (df["Period start time"].dt.date == start_date), selected_kpi]

        for date in unique_dates:
            ref_value = group.loc[group["Period start time"].dt.date == date, selected_kpi]

            if not orig_values.empty and not ref_value.empty:
                if comparison_operator == "> 'EX: DR'":
                    is_degraded = all(orig_value > ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((WCEL_name, orig_values, is_degraded, date, difference))
                elif comparison_operator == "< 'EX:CSSR'":
                    is_degraded = all(orig_value < ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((WCEL_name, orig_values.values[0], is_degraded, date, difference))

    compare_df = pd.DataFrame(compare, columns=["WCEL name", "Current value","Degraded", "dates", "difference"])
    difference_range = compare_df.groupby("WCEL name")["difference"].agg(["min", "max"]).reset_index()
    compare_df = compare_df.merge(difference_range, on="WCEL name")

    compare_df_pivot = compare_df.pivot(index=["WCEL name", "Current value"], columns="dates",
                                        values=["Degraded", "difference"])

    compare_df_pivot.to_excel(output_file_path)
    messagebox.showinfo("Success", "Excel file generated successfully.")


def safe_extract_first_element(x):
    if isinstance(x, pd.Series):
        if not x.empty and 0 in x.index:
            return x.iloc[0]
        else:
            return None  # Or some default value if the Series is empty or key does not exist
    return x


def generate_excel_LTE_LNCEL():
    input_file_path = input_file_path_entry.get()
    output_file_path = output_file_path_entry.get()
    selected_kpi = selected_kpi_var.get()
    start_date_str = entry_start_date.get()

    if not input_file_path or not output_file_path or not selected_kpi or not start_date_str:
        messagebox.showwarning("Warning", "Please select input file, output file, enter KPI, and select a start date.")
        return

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()

    df = pd.read_excel(input_file_path, index_col=False, sheet_name=0, engine='openpyxl')

    time_period = int(time_period_entry.get())

    if time_period == 14:
        # Get data for the same day of the week for the last 2 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 3)])]

    elif time_period == 30:
        # Get data for the same day of the week for the last 4 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 5)])]

    unique_dates = df2["Period start time"].dt.date.unique()

    print(unique_dates)

    columns_count = len(unique_dates)

    compare = []
    comparison_operator = comparison_operator_var.get()

    for LNCEL_name, group in df2.groupby("LNCEL name"):
        orig_values = df.loc[
            (df["LNCEL name"] == LNCEL_name) & (df["Period start time"].dt.date == start_date), selected_kpi]

        for date in unique_dates:
            ref_value = group.loc[group["Period start time"].dt.date == date, selected_kpi]
            print(date)
            if not orig_values.empty and not ref_value.empty:
                if comparison_operator == "> 'EX: DR'":
                    is_degraded = all(orig_value > ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    print(LNCEL_name)
                    compare.append((LNCEL_name, orig_values, is_degraded, date, difference))
                elif comparison_operator == "< 'EX:CSSR'":
                    is_degraded = all(orig_value < ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((LNCEL_name, orig_values.values[0], is_degraded, date, difference))

    compare_df = pd.DataFrame(compare, columns=["LNCEL name", "Current value","Degraded", "dates", "difference"])
    difference_range = compare_df.groupby("LNCEL name")["difference"].agg(["min", "max"]).reset_index()
    compare_df = compare_df.merge(difference_range, on="LNCEL name")

    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    print(compare_df.sample(10))

    compare_df_pivot = compare_df.pivot(index=["LNCEL name", "Current value"], columns="dates",
                                        values=["Degraded", "difference"])

    compare_df_pivot.to_excel(output_file_path)
    messagebox.showinfo("Success", "Excel file generated successfully.")

def generate_excel_LTE_LNBTS():
    input_file_path = input_file_path_entry.get()
    output_file_path = output_file_path_entry.get()
    selected_kpi = selected_kpi_var.get()
    start_date_str = entry_start_date.get()

    if not input_file_path or not output_file_path or not selected_kpi or not start_date_str:
        messagebox.showwarning("Warning", "Please select input file, output file, enter KPI, and select a start date.")
        return

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()

    df = pd.read_excel(input_file_path, index_col=False, sheet_name=0, engine='openpyxl')

    time_period = int(time_period_entry.get())

    if time_period == 14:
        # Get data for the same day of the week for the last 2 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 3)])]

    elif time_period == 30:
        # Get data for the same day of the week for the last 4 weeks
        df2 = df[df["Period start time"].dt.date.isin([start_date - timedelta(days=i*7) for i in range(1, 5)])]

    unique_dates = df2["Period start time"].dt.date.unique()
    columns_count = len(unique_dates)

    compare = []
    comparison_operator = comparison_operator_var.get()

    for LNBTS_name, group in df2.groupby("LNBTS name"):
        orig_values = df.loc[
            (df["LNBTS name"] == LNBTS_name) & (df["Period start time"].dt.date == start_date), selected_kpi]

        for date in unique_dates:
            ref_value = group.loc[group["Period start time"].dt.date == date, selected_kpi]

            if not orig_values.empty and not ref_value.empty:
                if comparison_operator == "> 'EX: DR'":
                    is_degraded = all(orig_value > ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((LNBTS_name, orig_values, is_degraded, date, difference))
                elif comparison_operator == "< 'EX:CSSR'":
                    is_degraded = all(orig_value < ref_value.values[0] for orig_value in orig_values.values)
                    difference = round(orig_values.values[0] - ref_value.values[0], 4) if is_degraded else np.nan
                    compare.append((LNBTS_name, orig_values.values[0], is_degraded, date, difference))

    compare_df = pd.DataFrame(compare, columns=["LNBTS name", "Current value","Degraded", "dates", "difference"])
    difference_range = compare_df.groupby("LNBTS name")["difference"].agg(["min", "max"]).reset_index()
    compare_df = compare_df.merge(difference_range, on="LNBTS name")

    compare_df_pivot = compare_df.pivot(index=["LNBTS name", "Current value"], columns="dates",
                                        values=["Degraded", "difference"])

    compare_df_pivot.to_excel(output_file_path)
    messagebox.showinfo("Success", "Excel file generated successfully.")

def generate_excel():
    if site_var.get() == 1:
        if "3G" in technology_var.get():
            generate_excel_3G_WBTS()
        else:
            generate_excel_LTE_LNBTS()
    elif cell_var.get() == 1:
        if "3G" in technology_var.get():
            generate_excel_3G_WCEL()
        else:
            generate_excel_LTE_LNCEL()
    else:
        messagebox.showwarning("Warning", "Please select either 'Site' or 'Cell'.")


root = Tk()
root.title("KPI Analysis")

Label(root, text="Input Excel File:").grid(row=0, column=0, padx=10, pady=5)
input_file_path_entry = Entry(root, width=50)
input_file_path_entry.grid(row=0, column=1, padx=10, pady=5)
Button(root, text="Browse", command=choose_input_file).grid(row=0, column=2, padx=10, pady=5)

Label(root, text="Output Excel File:").grid(row=1, column=0, padx=10, pady=5)
output_file_path_entry = Entry(root, width=50)
output_file_path_entry.grid(row=1, column=1, padx=10, pady=5)
Button(root, text="Browse", command=choose_output_file).grid(row=1, column=2, padx=10, pady=5)

Label(root, text="Time Period (14 or 30):").grid(row=2, column=0, padx=10, pady=5)
time_period_entry = Entry(root, width=10)
time_period_entry.grid(row=2, column=1, padx=10, pady=5)

Label(root, text="Start Date (YYYY-MM-DD):").grid(row=3, column=0, padx=10, pady=5)
entry_start_date = Entry(root, width=20)
entry_start_date.grid(row=3, column=1, padx=10, pady=5)

Label(root, text="KPI Name:").grid(row=4, column=0, padx=10, pady=5)
selected_kpi_var = StringVar()
if kpis_list:
    kpi_menu = OptionMenu(root, selected_kpi_var, *kpis_list)
else:
    kpi_menu = OptionMenu(root, selected_kpi_var, "No KPI selected")  # Default value
kpi_menu.grid(row=4, column=1, padx=10, pady=5)

Label(root, text="Comparison Operator:").grid(row=5, column=0, padx=10, pady=5)
comparison_operator_var = StringVar()
comparison_operator_var.set("< 'EX:CSSR'")  # Default value
comparison_operator_menu = OptionMenu(root, comparison_operator_var, "> 'EX: DR'", "< 'EX:CSSR'")
comparison_operator_menu.grid(row=5, column=1, padx=10, pady=5)

Label(root, text="Technology:").grid(row=6, column=0, padx=10, pady=5)
technology_var = StringVar()
technology_var.set("3G")  # Default value
technology_menu = OptionMenu(root, technology_var, "3G", "LTE")
technology_menu.grid(row=6, column=1, padx=10, pady=5)

site_var = IntVar()
cell_var = IntVar()

# Create checkboxes
site_checkbox = Checkbutton(root, text="Site", variable=site_var)
cell_checkbox = Checkbutton(root, text="Cell", variable=cell_var)

site_checkbox.grid(row=7, column=0, padx=10, pady=5)
cell_checkbox.grid(row=7, column=1, padx=10, pady=5)

Button(root, text="Generate Excel", command=generate_excel).grid(row=8, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()
