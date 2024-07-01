import pandas as pd
import matplotlib.pyplot as plt

# Function to create Excel file and plot graph
def create_excel_and_plot(data, output_filename):
    # Convert data to DataFrame
    df = pd.DataFrame(data, columns=['Action ID', 'Current (mA)', 'Time (sec)'])

    # Create Excel file with .xlsx extension
    output_filename = output_filename if output_filename.endswith('.xlsx') else f"{output_filename}.xlsx"
    df.to_excel(output_filename, index=False, engine='openpyxl')

    # Plot step graph
    x_values = []
    y_values = []
    time = 0 
    for i in range(len(df)):
        if i == 0:
            x_values.extend([0, df['Time (sec)'].iloc[i]])
            y_values.extend([df['Current (mA)'].iloc[i], df['Current (mA)'].iloc[i]])
        else:
            x_values.extend([time, (df['Time (sec)'].iloc[i]+time)])
            y_values.extend([df['Current (mA)'].iloc[i], df['Current (mA)'].iloc[i]])
        time = time + df['Time (sec)'].iloc[i]

    plt.plot(x_values, y_values, label='Current Consumption', linestyle='-', marker='o')
    plt.xlabel('Time (sec)')
    plt.ylabel('Current (mA)')
    plt.title('Current Consumption Over Time')
    plt.legend()
    plt.savefig(f"{output_filename.replace('.xlsx', '_step_graph.png')}")
    plt.show()

# Main function
def main():
    actions = {}
    data = []
    action_id_counter = 1

    while True:
        print("Options:")
        print("0 - Insert manually")
        print("1 - Use a saved action")
        print("END - Finish input")
        option = input("Enter option: ")

        if option == '0':
            action = input("Enter action name: ")
            if action.upper() == 'END':
                output_filename = input("Enter Excel file name: ")
                create_excel_and_plot(data, output_filename)
                break

            # Check if the action is already in the dictionary
            if action not in actions:
                actions[action] = action_id_counter
                current = float(input("Enter current consumption (mA): "))
                time = float(input("Enter time (sec): "))
                action_id = action_id_counter
                action_id_counter += 1
            else:
                action_id = actions[action]
                current = 0  # Default current for existing action
                time = action_id_counter  # Action ID counter for existing action

            data.append([action_id, current, time])

        elif option == '1':
            print("Saved Actions:")
            for idx, (action_name, action_id) in enumerate(actions.items(), start=1):
                print(f"{idx} - {action_name}")

            selected_number = int(input("Enter the number of the saved action: "))
            if 1 <= selected_number <= len(actions):
                selected_action = list(actions.keys())[selected_number - 1]
                action_id = actions[selected_action]
                current = data[selected_number-1][1]  # Default current for existing action
                time = data[selected_number-1][2] # Action ID counter for existing action
                data.append([action_id, current, time])
            else:
                print("Invalid number. Please choose a valid number.")

        elif option.upper() == 'END':
            output_filename = input("Enter Excel file name: ")
            create_excel_and_plot(data, output_filename)
            break
        else:
            print("Invalid option. Please enter 0, 1, or END.")

if __name__ == "__main__":
    main()
