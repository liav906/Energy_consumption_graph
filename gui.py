import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class CurrentConsumptionApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Current Consumption App")

        self.actions = {}
        self.data = []
        self.action_id_counter = 1

        self.create_widgets()

    def create_widgets(self):
        self.manual_button = ttk.Button(self.master, text="Manually", command=self.insert_manually)
        self.manual_button.grid(row=0, column=0, pady=5)

        self.saved_button = ttk.Button(self.master, text="Saved", command=self.use_saved_action)
        self.saved_button.grid(row=0, column=1, pady=5)

        self.done_button = ttk.Button(self.master, text="Done", command=self.finish_input)
        self.done_button.grid(row=0, column=2, pady=5)

        self.output_label = ttk.Label(self.master, text="Enter Excel file name:")
        self.output_label.grid(row=1, column=0, columnspan=3)

        self.output_filename_entry = ttk.Entry(self.master)
        self.output_filename_entry.grid(row=2, column=0, columnspan=3, pady=5)

    def insert_manually(self):
        manual_window = tk.Toplevel(self.master)
        manual_window.title("Manual Entry")

        action_label = ttk.Label(manual_window, text="Enter action name:")
        action_label.grid(row=0, column=0, pady=5)
        action_entry = ttk.Entry(manual_window)
        action_entry.grid(row=0, column=1, pady=5)

        current_label = ttk.Label(manual_window, text="Enter current consumption (mA):")
        current_label.grid(row=1, column=0, pady=5)
        current_entry = ttk.Entry(manual_window)
        current_entry.grid(row=1, column=1, pady=5)

        time_label = ttk.Label(manual_window, text="Enter time (sec):")
        time_label.grid(row=2, column=0, pady=5)
        time_entry = ttk.Entry(manual_window)
        time_entry.grid(row=2, column=1, pady=5)

        confirm_button = ttk.Button(manual_window, text="Confirm", command=lambda: self.confirm_manual_entry(
            action_entry.get(), float(current_entry.get()), float(time_entry.get()), manual_window))
        confirm_button.grid(row=3, column=0, columnspan=2, pady=5)

    def confirm_manual_entry(self, action, current, time, window):
        action_id = self.action_id_counter
        self.action_id_counter += 1
        self.data.append([action_id, current, time])
        self.actions[action] = action_id
        window.destroy()

    def use_saved_action(self):
        saved_window = tk.Toplevel(self.master)
        saved_window.title("Use Saved Action")

        for idx, (action_name, action_id) in enumerate(self.actions.items(), start=1):
            button = ttk.Button(saved_window, text=action_name, command=lambda name=action_name: self.confirm_saved_action(name, saved_window))
            button.grid(row=idx - 1, column=0, pady=5)

    def confirm_saved_action(self, action, window):
        action_id = self.actions[action]
        current = self.data[action_id - 1][1]
        time = self.data[action_id - 1][2]
        self.data.append([action_id, current, time])
        window.destroy()

    def finish_input(self):
        output_filename = self.output_filename_entry.get()
        self.create_excel_and_plot(output_filename)

    def create_excel_and_plot(self, output_filename):
        df = pd.DataFrame(self.data, columns=['Action ID', 'Current (mA)', 'Time (sec)'])
        output_filename = output_filename if output_filename.endswith('.xlsx') else f"{output_filename}.xlsx"
        df.to_excel(output_filename, index=False, engine='openpyxl')

        x_values = []
        y_values = []
        time = 0
        for i in range(len(df)):
            if i == 0:
                x_values.extend([0, df['Time (sec)'].iloc[i]])
                y_values.extend([df['Current (mA)'].iloc[i], df['Current (mA)'].iloc[i]])
            else:
                x_values.extend([time, (df['Time (sec)'].iloc[i] + time)])
                y_values.extend([df['Current (mA)'].iloc[i], df['Current (mA)'].iloc[i]])
            time = time + df['Time (sec)'].iloc[i]

        fig, ax = plt.subplots()
        ax.plot(x_values, y_values, label='Current Consumption', linestyle='-', marker='o')
        ax.set_xlabel('Time (sec)')
        ax.set_ylabel('Current (mA)')
        ax.set_title('Current Consumption Over Time')
        ax.legend()
        plt.savefig(f"{output_filename.replace('.xlsx', '_step_graph.png')}")
        canvas = FigureCanvasTkAgg(fig, master=self.master)
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.grid(row=3, column=0, columnspan=3, pady=5)

def main():
    root = tk.Tk()
    app = CurrentConsumptionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()