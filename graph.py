import pandas as pd
import matplotlib.pyplot as plt
from collections import defaultdict

# Load the data from Excel
file_name = 'QDS-above-70-crossed-40d.xlsx'
sheet_name = 'QDS-above-70-crossed-40d'

# Load the data into a pandas DataFrame
df = pd.read_excel(file_name, sheet_name=sheet_name)

# Column names
machine_col = 'NetBIOS'
application_col = 'Application'

# Step 1: Group by application and list affected machines
application_groups = df.groupby(application_col)[machine_col].apply(set)

# Step 2: Count the number of unique machines each application fixes
application_machine_count = {app: len(machines) for app, machines in application_groups.items()}

# Step 3: Sort applications by the number of machines they fix (descending)
sorted_applications = sorted(application_machine_count.items(), key=lambda x: x[1], reverse=True)

# Step 4: Greedy selection of applications to patch
target_machines = 140
selected_applications = []
fixed_machines = set()

for app, count in sorted_applications:
    # Get machines fixed by this application
    machines_fixed_by_app = application_groups[app]
    
    # Add these machines to the fixed list
    fixed_machines.update(machines_fixed_by_app)
    selected_applications.append((app, count))
    
    # Stop if the target number of machines is fixed
    if len(fixed_machines) >= target_machines:
        break

# Step 5: Prepare Results DataFrame
result_df = pd.DataFrame(selected_applications, columns=["Application", "Machines Fixed by Application"])
result_df["Cumulative Machines Fixed"] = result_df["Machines Fixed by Application"].cumsum()

# Step 6: Generate Pivot Table
pivot_table = df.groupby(application_col)[machine_col].nunique().reset_index()
pivot_table = pivot_table.rename(columns={machine_col: 'Unique Machines Affected'})

# Step 7: Save all data to a single Excel file
output_file = "patched_machines_analysis.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Original data
    df.to_excel(writer, sheet_name="Original Data", index=False)
    # Pivot table
    pivot_table.to_excel(writer, sheet_name="Pivot Table", index=False)
    # Selected applications
    result_df.to_excel(writer, sheet_name="Selected Applications", index=False)

# Step 8: Plot the bar graph
plt.figure(figsize=(12, 6))
plt.barh(result_df["Application"], result_df["Machines Fixed by Application"], color='skyblue')
plt.xlabel('Number of Machines Fixed')
plt.ylabel('Applications')
plt.title('Applications Selected to Fix Machines')
plt.gca().invert_yaxis()  # Invert y-axis for better readability
plt.tight_layout()

# Save the graph
graph_file = 'application_vs_machines_fixed.png'
plt.savefig(graph_file)
plt.show()

# Print Summary
print(f"\nApplications to Patch to Fix Target Machines:")
print(result_df)
print(f"\nTotal Machines Fixed: {len(fixed_machines)}")
print(f"Target Machines: {target_machines}")
print(f"Results saved in: {output_file}")
print(f"Graph saved as: {graph_file}")
