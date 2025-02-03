import pyodbc
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import os
import datetime

# Path to your Access database file
database_path = r'\\srvfileshare\Departments\9360 - Information Technology\Databases\DesktopInventory\DesktopInventory.accdb'

# Connection string
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=' + database_path + ';'
)

# Establish connection
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# Functions to interact with the database
def add_item(name, description, category):
    cursor.execute("INSERT INTO Items (Name, Description, Category) VALUES (?, ?, ?)", (name, description, category))
    conn.commit()
    cursor.execute("SELECT @@IDENTITY AS NewItemID")
    item_id = cursor.fetchone().NewItemID
    initialize_inventory(item_id, 0, 0, False)  # Initialize OnOrder as False
    return item_id

def remove_item(item_id):
    # Update Status column to 'Disabled' instead of deleting
    cursor.execute("UPDATE Items SET Status = 'Disabled' WHERE ItemID = ?", (item_id,))
    conn.commit()

def update_inventory(item_id, quantity):
    cursor.execute("UPDATE Inventory SET Quantity = ? WHERE ItemID = ?", (quantity, item_id))
    conn.commit()
    refresh_all_inventories()

def set_threshold(item_id, threshold):
    cursor.execute("UPDATE Inventory SET Threshold = ? WHERE ItemID = ?", (threshold, item_id))
    conn.commit()
    refresh_all_inventories()

def log_transaction(item_id, quantity, transaction_type, sr=None):
    user = os.getlogin()
    if sr:
        cursor.execute("INSERT INTO Transactions (ItemID, Quantity, TransactionType, [User], SR) VALUES (?, ?, ?, ?, ?)",
                       (item_id, quantity, transaction_type, user, sr))
    else:
        cursor.execute("INSERT INTO Transactions (ItemID, Quantity, TransactionType, [User]) VALUES (?, ?, ?, ?)",
                       (item_id, quantity, transaction_type, user))
    conn.commit()


import tkinter as tk
from tkinter import font as tkfont  # Import the font module

# Other imports and connection setup

def refresh_inventory(tree, category):
    # Clear existing items in the Treeview
    for item in tree.get_children():
        tree.delete(item)
    
    # Retrieve data from database and populate Treeview
    cursor.execute("SELECT Items.Name, Inventory.Quantity, Inventory.Purchased, Inventory.Threshold FROM Items INNER JOIN Inventory ON Items.ItemID = Inventory.ItemID WHERE Items.Category = ? AND Items.Status = 'Active'", (category,))
    
    for row in cursor.fetchall():
        name = row.Name
        quantity = row.Quantity
        purchased = row.Purchased
        threshold = row.Threshold
        
        # Determine background color based on conditions
        if quantity < threshold:
            if purchased:
                color = "green"
            else:
                color = "red"
        else:
            color = "white"  # Default background color
        
        tree.insert('', 'end', values=(name, quantity, purchased, threshold), tags=(name,))
        tree.tag_configure(name, background=color)
    
    # Auto adjust column widths
    font = tkfont.Font()
    for col in tree["columns"]:
        # Measure the width of the column header
        max_width = font.measure(col)
        for item in tree.get_children():
            # Measure the width of the item text in this column
            item_text = tree.item(item, "values")[tree["columns"].index(col)]
            max_width = max(max_width, font.measure(item_text))
        tree.column(col, width=max_width + 10)  # Add a little extra padding


def initialize_inventory(item_id, initial_quantity, threshold, purchased=False):
    cursor.execute("INSERT INTO Inventory (ItemID, Quantity, Threshold, Purchased) VALUES (?, ?, ?, ?)", (item_id, initial_quantity, threshold, purchased))
    conn.commit()

# Admin Mode Functions
admin_mode_active = False

def admin_mode():
    global admin_mode_active
    password = simpledialog.askstring("Admin Mode", "Enter admin password:", show='*')
    # Check if password is correct (dummy check, replace with actual logic)
    if password == "Malachi":
        messagebox.showinfo("Admin Mode", "Admin Mode activated.")
        admin_mode_active = True
        admin_menu.entryconfig("Admin Mode", state="disabled")  # Disable the Admin Mode option after successful entry
    else:
        messagebox.showerror("Admin Mode", "Incorrect password")

def item_purchased():
    selected_item = get_selected_item()
    if not selected_item:
        messagebox.showerror("Error", "No item selected")
        return

    item_name = selected_item["values"][0]
    print(f"Selected item name: {item_name}")

    # Fetch the ItemID using the item name
    cursor.execute("SELECT ItemID FROM Items WHERE Name = ?", (item_name,))
    item_id_result = cursor.fetchone()
    
    if item_id_result:
        item_id = item_id_result.ItemID
        print(f"Fetched item ID: {item_id}")
    else:
        messagebox.showerror("Error", "Item not found in database.")
        return

    # Create a dialog to get the quantity and SR number
    purchase_dialog = tk.Toplevel(app)
    purchase_dialog.title("Item Purchased")

    tk.Label(purchase_dialog, text="Quantity of item purchased:").grid(row=0, column=0, padx=10, pady=5)
    quantity_entry = tk.Entry(purchase_dialog)
    quantity_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(purchase_dialog, text="SR Number:").grid(row=1, column=0, padx=10, pady=5)
    sr_entry = tk.Entry(purchase_dialog)
    sr_entry.grid(row=1, column=1, padx=10, pady=5)

    def save_purchase():
        try:
            quantity_to_purchase = int(quantity_entry.get().strip())
            sr_number = sr_entry.get().strip()

            if quantity_to_purchase <= 0:
                messagebox.showerror("Error", "Please enter a valid positive quantity.")
                return

            # Retrieve the current 'Purchased' value from the database
            cursor.execute("SELECT Purchased FROM Inventory WHERE ItemID = ?", (item_id,))
            result = cursor.fetchone()
            
            print(f"Database query result: {result}")

            if result:
                current_purchased = result.Purchased
                current_purchased = int(current_purchased) if current_purchased is not None else 0
            else:
                messagebox.showerror("Error", "Item not found in inventory.")
                return

            # Increment the 'Purchased' value by the quantity to purchase
            new_purchased = current_purchased + quantity_to_purchase

            # Update the inventory with the new 'Purchased' value
            cursor.execute("UPDATE Inventory SET Purchased = ? WHERE ItemID = ?", (new_purchased, item_id))
            conn.commit()

            # Log the transaction
            log_transaction(item_id, quantity_to_purchase, "Purchased", sr_number)

            # Refresh inventory views to reflect changes
            refresh_all_inventories()

            # Close the dialog
            purchase_dialog.destroy()

        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for quantity.")

    tk.Button(purchase_dialog, text="Save", command=save_purchase).grid(row=2, columnspan=2, pady=10)

    purchase_dialog.transient(app)  # Set the dialog transient to the main app
    purchase_dialog.grab_set()  # Grab focus to the dialog
    app.wait_window(purchase_dialog)  # Wait for the dialog to be closed before returning

def item_deployed():
    selected_item = get_selected_item()
    if not selected_item:
        messagebox.showerror("Error", "No item selected")
        return

    item_name = selected_item["values"][0]
    print(f"Selected item name: {item_name}")

    # Fetch the ItemID using the item name
    cursor.execute("SELECT ItemID FROM Items WHERE Name = ?", (item_name,))
    item_id_result = cursor.fetchone()

    if item_id_result:
        item_id = item_id_result.ItemID
        print(f"Fetched item ID: {item_id}")
    else:
        messagebox.showerror("Error", "Item not found in database.")
        return

    # Create a dialog to get the quantity and SR number
    deploy_dialog = tk.Toplevel(app)
    deploy_dialog.title("Deploy Item")

    tk.Label(deploy_dialog, text="Quantity of item deployed:").grid(row=0, column=0, padx=10, pady=5)
    quantity_entry = tk.Entry(deploy_dialog)
    quantity_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(deploy_dialog, text="SR Number:").grid(row=1, column=0, padx=10, pady=5)
    sr_entry = tk.Entry(deploy_dialog)
    sr_entry.grid(row=1, column=1, padx=10, pady=5)

    def save_deploy():
        try:
            quantity_to_deploy = int(quantity_entry.get().strip())
            sr_number = sr_entry.get().strip()

            if quantity_to_deploy <= 0:
                messagebox.showerror("Error", "Please enter a valid positive quantity.")
                return

            # Retrieve the current 'Quantity' value from the database
            cursor.execute("SELECT Quantity FROM Inventory WHERE ItemID = ?", (item_id,))
            result = cursor.fetchone()

            if result:
                current_quantity = result.Quantity
                current_quantity = int(current_quantity) if current_quantity is not None else 0
            else:
                messagebox.showerror("Error", "Item not found in inventory.")
                return

            if quantity_to_deploy > current_quantity:
                messagebox.showerror("Error", "Quantity to deploy exceeds current quantity.")
                return

            # Update the 'Quantity' value
            new_quantity = current_quantity - quantity_to_deploy

            # Update the inventory with the new 'Quantity' value
            cursor.execute("UPDATE Inventory SET Quantity = ? WHERE ItemID = ?", (new_quantity, item_id))
            conn.commit()

            # Log the deployment transaction
            log_transaction(item_id, quantity_to_deploy, "Deployed", sr_number)

            # Refresh inventory views to reflect changes
            refresh_all_inventories()

            # Close the dialog
            deploy_dialog.destroy()

        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for quantity.")

    tk.Button(deploy_dialog, text="Save", command=save_deploy).grid(row=2, columnspan=2, pady=10)

    deploy_dialog.transient(app)  # Set the dialog transient to the main app
    deploy_dialog.grab_set()  # Grab focus to the dialog
    app.wait_window(deploy_dialog)  # Wait for the dialog to be closed before returning






def item_received():
    selected_item = get_selected_item()
    if not selected_item:
        messagebox.showerror("Error", "No item selected")
        return

    item_id = get_item_id(selected_item)

    # Create a dialog to get the number received and SR number
    receive_dialog = tk.Toplevel(app)
    receive_dialog.title("Item Received")

    tk.Label(receive_dialog, text="Number of items received:").grid(row=0, column=0, padx=10, pady=5)
    number_received_entry = tk.Entry(receive_dialog)
    number_received_entry.grid(row=0, column=1, padx=10, pady=5)

    tk.Label(receive_dialog, text="SR Number:").grid(row=1, column=0, padx=10, pady=5)
    sr_entry = tk.Entry(receive_dialog)
    sr_entry.grid(row=1, column=1, padx=10, pady=5)

    def save_received():
        try:
            number_received = int(number_received_entry.get().strip())
            sr_number = sr_entry.get().strip()

            if number_received <= 0:
                messagebox.showerror("Error", "Please enter a valid positive number.")
                return

            # Retrieve the current 'Purchased' and 'Quantity' values from the database
            cursor.execute("SELECT Purchased, Quantity FROM Inventory WHERE ItemID = ?", (item_id,))
            result = cursor.fetchone()
            if result:
                current_purchased = result.Purchased
                current_quantity = result.Quantity
                
                # Ensure values are integers
                current_purchased = int(current_purchased) if current_purchased is not None else 0
                current_quantity = int(current_quantity) if current_quantity is not None else 0
            else:
                messagebox.showerror("Error", "Item not found in inventory.")
                return

            if number_received > current_purchased:
                messagebox.showerror("Error", "Number received exceeds the number purchased.")
                return

            # Update the 'Purchased' and 'Quantity' values
            new_purchased = current_purchased - number_received
            new_quantity = current_quantity + number_received

            # Update the inventory with the new values
            cursor.execute("UPDATE Inventory SET Purchased = ?, Quantity = ? WHERE ItemID = ?", (new_purchased, new_quantity, item_id))
            conn.commit()

            # Log the transaction
            log_transaction(item_id, number_received, "Received", sr_number)

            # Refresh inventory views to reflect changes
            refresh_all_inventories()

            # Close the dialog
            receive_dialog.destroy()

        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for quantity.")

    tk.Button(receive_dialog, text="Save", command=save_received).grid(row=2, columnspan=2, pady=10)

    receive_dialog.transient(app)  # Set the dialog transient to the main app
    receive_dialog.grab_set()  # Grab focus to the dialog
    app.wait_window(receive_dialog)  # Wait for the dialog to be closed before returning


# Report Functions
def generate_needed_items_report():
    report_window = tk.Toplevel(app)
    report_window.title("Needed Items Report")

    report_text = ""

    cursor.execute("SELECT Items.Name, Inventory.Quantity, Inventory.Threshold, Inventory.Purchased FROM Items INNER JOIN Inventory ON Items.ItemID = Inventory.ItemID")
    for row in cursor.fetchall():
        if row.Quantity < row.Threshold and not row.Purchased:
            report_text += f"{row.Name}: Needs to purchase {row.Threshold - row.Quantity} units\n"
        elif row.Quantity < row.Threshold and row.Purchased:
            report_text += f"{row.Name}: Already on order, awaiting delivery\n"

    if report_text:
        report_label = tk.Label(report_window, text=report_text, padx=10, pady=10)
        report_label.pack()
    else:
        no_items_label = tk.Label(report_window, text="No items currently need to be purchased.")
        no_items_label.pack()

# GUI-related functions
def add_item_ui():
    def save_item():
        name = name_entry.get()
        description = description_entry.get()
        category = category_entry.get()
        add_item(name, description, category)
        refresh_all_inventories()
        add_window.destroy()

    add_window = tk.Toplevel(app)
    add_window.title("Add Item")

    tk.Label(add_window, text="Name").grid(row=0, column=0)
    tk.Label(add_window, text="Description").grid(row=1, column=0)
    tk.Label(add_window, text="Category").grid(row=2, column=0)

    name_entry = tk.Entry(add_window)
    description_entry = tk.Entry(add_window)
    category_entry = tk.Entry(add_window)

    name_entry.grid(row=0, column=1)
    description_entry.grid(row=1, column=1)
    category_entry.grid(row=2, column=1)

    tk.Button(add_window, text="Save", command=save_item).grid(row=3, columnspan=2, pady=10)

def edit_item_ui():
    selected_item = get_selected_item()
    if not selected_item:
        messagebox.showerror("Error", "No item selected")
        return

    item_id = get_item_id(selected_item)

    def save_item():
        name = name_entry.get()
        description = description_entry.get()
        category = category_entry.get()
        quantity = int(quantity_entry.get())
        threshold = int(threshold_entry.get())
        purchased = purchased_var.get()
        
        cursor.execute("UPDATE Items SET Name = ?, Description = ?, Category = ? WHERE ItemID = ?", (name, description, category, item_id))
        cursor.execute("UPDATE Inventory SET Quantity = ?, Threshold = ?, Purchased = ? WHERE ItemID = ?", (quantity, threshold, purchased, item_id))
        conn.commit()
        refresh_all_inventories()
        edit_window.destroy()

    cursor.execute("SELECT Items.Name, Items.Description, Items.Category, Inventory.Quantity, Inventory.Threshold, Inventory.Purchased FROM Items INNER JOIN Inventory ON Items.ItemID = Inventory.ItemID WHERE Items.ItemID = ?", (item_id,))
    item = cursor.fetchone()

    edit_window = tk.Toplevel(app)
    edit_window.title("Edit Item")

    tk.Label(edit_window, text="Name").grid(row=0, column=0)
    tk.Label(edit_window, text="Description").grid(row=1, column=0)
    tk.Label(edit_window, text="Category").grid(row=2, column=0)
    tk.Label(edit_window, text="Quantity").grid(row=3, column=0)
    tk.Label(edit_window, text="Threshold").grid(row=4, column=0)
    tk.Label(edit_window, text="On Order").grid(row=5, column=0)

    name_entry = tk.Entry(edit_window)
    name_entry.insert(0, item.Name)
    description_entry = tk.Entry(edit_window)
    description_entry.insert(0, item.Description)
    category_entry = tk.Entry(edit_window)
    category_entry.insert(0, item.Category)
    quantity_entry = tk.Entry(edit_window)
    quantity_entry.insert(0, item.Quantity)
    threshold_entry = tk.Entry(edit_window)
    threshold_entry.insert(0, item.Threshold)
    purchased_var = tk.BooleanVar(value=item.Purchased)
    purchased_check = tk.Checkbutton(edit_window, variable=purchased_var)

    name_entry.grid(row=0, column=1)
    description_entry.grid(row=1, column=1)
    category_entry.grid(row=2, column=1)
    quantity_entry.grid(row=3, column=1)
    threshold_entry.grid(row=4, column=1)
    purchased_check.grid(row=5, column=1)

    tk.Button(edit_window, text="Save", command=save_item).grid(row=6, columnspan=2, pady=10)

def remove_item_ui():
    selected_item = get_selected_item()
    if not selected_item:
        messagebox.showerror("Error", "No item selected")
        return

    item_id = get_item_id(selected_item)
    remove_item(item_id)
    refresh_all_inventories()

def get_item_id(item):
    cursor.execute("SELECT ItemID FROM Items WHERE Name = ?", (item['values'][0],))
    return cursor.fetchone().ItemID

def get_selected_item():
    selected = major_tree.selection() or minor_tree.selection()
    if not selected:
        return None
    item = major_tree.item(selected) if major_tree.selection() else minor_tree.item(selected)
    return item

def show_deployment_report_options():
    # Create a new window for selecting items and date range
    report_window = tk.Toplevel()
    report_window.title("Major Item Deployment Report")
    
    tk.Label(report_window, text="Select Items:").grid(row=0, column=0, padx=10, pady=10)
    
    # List of items (you can load this from the database or use predefined items)
    items_listbox = tk.Listbox(report_window, selectmode=tk.MULTIPLE)
    items_listbox.grid(row=1, column=0, padx=10, pady=10)
    items = get_inventory_items()  # Function to get available items
    for item in items:
        items_listbox.insert(tk.END, item)
    
    tk.Label(report_window, text="Start Date:").grid(row=2, column=0, padx=10, pady=10)
    start_date_entry = tk.Entry(report_window)
    start_date_entry.grid(row=3, column=0, padx=10, pady=10)
    start_date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
    
    tk.Label(report_window, text="End Date:").grid(row=4, column=0, padx=10, pady=10)
    end_date_entry = tk.Entry(report_window)
    end_date_entry.grid(row=5, column=0, padx=10, pady=10)
    end_date_entry.insert(0, datetime.datetime.now().strftime("%Y-%m-%d"))
    
    def generate_report():
        selected_items = [items_listbox.get(i) for i in items_listbox.curselection()]
        start_date = start_date_entry.get()
        end_date = end_date_entry.get()

        if not selected_items:
            messagebox.showerror("Error", "Please select at least one item.")
            return
        
        if not start_date or not end_date:
            messagebox.showerror("Error", "Please enter valid start and end dates.")
            return
        
        # Call the function to generate the report based on selected items and date range
        generate_major_item_deployment_report(selected_items, start_date, end_date)
        report_window.destroy()

    tk.Button(report_window, text="Generate Report", command=generate_report).grid(row=6, column=0, padx=10, pady=10)

def generate_major_item_deployment_report():
    try:
        # Define the query
        query = """
            SELECT ItemName, SUM(QuantityDeployed) AS TotalDeployed
            FROM ItemDeployment
            GROUP BY ItemName
            HAVING SUM(QuantityDeployed) > 0
            ORDER BY TotalDeployed DESC
        """
        
        # Execute the query using the existing cursor
        cursor.execute(query)
        result = cursor.fetchall()

        # Process the result and display in a report (example code)
        report_data = "\n".join([f"Item: {item[0]}, Total Deployed: {item[1]}" for item in result])

        # Example of how to display the report (you can customize this)
        messagebox.showinfo("Deployment Report", report_data)
        
    except Exception as e:
        messagebox.showerror("Error", f"Error generating report: {e}")


def get_inventory_items():
    try:
        # Use the existing connection (conn) and cursor
        cursor.execute("SELECT DISTINCT Name FROM Items")
        items = cursor.fetchall()
        return [item[0] for item in items]
    except Exception as e:
        messagebox.showerror("Error", f"Error fetching items: {e}")
        return []





def refresh_all_inventories():
    refresh_inventory(major_tree, 'Major')
    refresh_inventory(minor_tree, 'Minor')

# Function to show About dialog
def show_about_dialog():
    about_text = """
    Inventory Management System
    Version 1.2.1
    
    Change Log:
    Version 1.2.1
    - Fixed bugs related to the Deploy Function

    Version 1.2.0
    - Made changes to several admin functions
    - Added Purchased fucntion
    - Added SR requirement
    - Added On Order Field

    Version 1.1.0
    - Added SR requirement for Major Items
    - Bug Fixes

    Version 1.0.0
    - Initial release
    """
    messagebox.showinfo("About", about_text)

app = tk.Tk()
app.title("Inventory Management System")

# Menu Bar
menu_bar = tk.Menu(app)

# Main Menu
main_menu = tk.Menu(menu_bar, tearoff=0)

#About Menu
about_menu = tk.Menu(menu_bar,tearoff=0)
about_menu.add_command(label="About", command=show_about_dialog)
main_menu.add_cascade(label="About", menu=about_menu)

# Admin Menu
admin_menu = tk.Menu(main_menu, tearoff=0)
admin_menu.add_command(label="Admin Mode", command=admin_mode)
main_menu.add_cascade(label="Admin", menu=admin_menu)

# Report Menu
report_menu = tk.Menu(main_menu, tearoff=0)
report_menu.add_command(label="Generate Needed Items Report", command=generate_needed_items_report)
report_menu.add_command(label="Generate Item Deployment Report", command=show_deployment_report_options)
main_menu.add_cascade(label="Reports", menu=report_menu)

menu_bar.add_cascade(label="Options", menu=main_menu)

app.config(menu=menu_bar)


# Notebook for tabs
notebook = ttk.Notebook(app)
notebook.pack(fill='both', expand=True)

# Major Items Tab
major_frame = ttk.Frame(notebook)
notebook.add(major_frame, text="Major Items")

columns = ("Name", "Quantity", "On Order", "Threshold")

major_tree = ttk.Treeview(major_frame, columns=columns, show="headings")
for col in columns:
    major_tree.heading(col, text=col)
    major_tree.column(col, width=100)
major_tree.pack(fill="both", expand=True)

# Minor Items Tab
minor_frame = ttk.Frame(notebook)
notebook.add(minor_frame, text="Minor Items")

minor_tree = ttk.Treeview(minor_frame, columns=columns, show="headings")
for col in columns:
    minor_tree.heading(col, text=col)
    minor_tree.column(col, width=100)
minor_tree.pack(fill="both", expand=True)

refresh_all_inventories()

# Right-click menu
def show_context_menu(event):
    selected_item = get_selected_item()
    if not selected_item:
        return

    context_menu = tk.Menu(app, tearoff=0)
    context_menu.add_command(label="Item Purchased", command=item_purchased)
    context_menu.add_command(label="Item Received", command=item_received)
    context_menu.add_command(label="Deploy Item", command=item_deployed)  # Added Deploy Item option

    # Only show these options if admin mode is active
    if admin_mode_active:
        context_menu.add_command(label="Add Item", command=add_item_ui)
        context_menu.add_command(label="Edit", command=edit_item_ui)
        context_menu.add_command(label="Delete", command=remove_item_ui)

    context_menu.tk_popup(event.x_root, event.y_root)

# Bind the context menu to the Treeview widgets
major_tree.bind("<Button-3>", show_context_menu)
minor_tree.bind("<Button-3>", show_context_menu)


major_tree.bind("<Button-3>", show_context_menu)
minor_tree.bind("<Button-3>", show_context_menu)

app.mainloop()

# Close the database connection when the application is closed
conn.close()
