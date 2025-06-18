import os
import pandas as pd

from tabulate import tabulate

FILE_PATH = "/Users/mohammedshabeel/Desktop/python/employees.xlsx"
COLUMNS = ["ID", "Name", "Age", "Department"]

def load_data():
    # Check if file exists
    if not os.path.exists(FILE_PATH):
        print("📁 Excel file not found. Creating a new one...")
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(FILE_PATH, index=False, engine='openpyxl')
        print("✅ New Excel file created.")
        return df

    # Try reading the file if it exists
    try:
        df = pd.read_excel(FILE_PATH, engine='openpyxl')

        # Validate structure
        if df.empty or not all(col in df.columns for col in COLUMNS):
            raise ValueError("Invalid or empty Excel file structure.")

        print("✅ Excel file loaded successfully.")
        return df
    except Exception as e:
        print(f"⚠️ Error loading Excel file: {e}")
        print("🔄 Recreating file...")

        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(FILE_PATH, index=False, engine='openpyxl')
        print("✅ New Excel file created.")
        return df
# Save to Excel
def save_data(df):
    df.to_excel(FILE_PATH, index=False)
    print("✅ Data saved to Excel.\n")

# Display all employees
def view_employees(df):
    if df.empty:
        print("⚠️ No employee records found.\n")
    else:
        print("\n📋 Employee Records:")
        print(tabulate(df, headers = 'keys', tablefmt ='fancy_grid'))
        #print(df.to_string(index=False), "\n")

# Add a new employee
def add_employee(df):
    try:
        new_id = int(input("Enter ID: "))
        if new_id in df["ID"].values:
            print("❌ ID already exists!\n")
            return df

        name = input("Enter Name: ")
        age = int(input("Enter Age: "))
        dept = input("Enter Department: ")

        new_row = pd.DataFrame([{
            "ID": new_id,
            "Name": name,
            "Age": age,
            "Department": dept
        }])

        df = pd.concat([df, new_row], ignore_index=True)
        print("✅ Employee added.\n")
    except Exception as e:
        print("❌ Error adding employee:", e, "\n")
    return df

# Update employee
def update_employee(df):
    try:
        emp_id = int(input("Enter ID to update: "))
        if emp_id not in df["ID"].values:
            print("❌ ID not found.\n")
            return df

        index = df.index[df["ID"] == emp_id][0]
        print("Leave blank to keep existing value.")

        name = input(f"New Name [{df.at[index, 'Name']}]: ") or df.at[index, "Name"]
        age_input = input(f"New Age [{df.at[index, 'Age']}]: ")
        age = int(age_input) if age_input else df.at[index, "Age"]
        dept = input(f"New Department [{df.at[index, 'Department']}]: ") or df.at[index, "Department"]

        df.at[index, "Name"] = name
        df.at[index, "Age"] = age
        df.at[index, "Department"] = dept

        print("✅ Employee updated.\n")
    except Exception as e:
        print("❌ Error updating employee:", e, "\n")
    return df

# Delete employee
def delete_employee(df):
    try:
        emp_id = int(input("Enter ID to delete: "))
        if emp_id not in df["ID"].values:
            print("❌ ID not found.\n")
            return df

        df = df[df["ID"] != emp_id]
        print("✅ Employee deleted.\n")
    except Exception as e:
        print("❌ Error deleting employee:", e, "\n")
    return df

# Main menu
def main():
    df = load_data()

    while True:
        print("📁 Employee Manager")
        print("1. View Employees")
        print("2. Add Employee")
        print("3. Update Employee")
        print("4. Delete Employee")
        print("5. Save & Exit")

        choice = input("Enter choice: ")

        if choice == '1':
            view_employees(df)
        elif choice == '2':
            df = add_employee(df)
        elif choice == '3':
            df = update_employee(df)
        elif choice == '4':
            df = delete_employee(df)
        elif choice == '5':
            save_data(df)
            break
        else:
            print("❌ Invalid choice. Try again.\n")

if __name__ == "__main__":
    main()
