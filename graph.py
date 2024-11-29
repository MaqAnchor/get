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

# Step 1: Create a pivot table
pivot_table = df.groupby(application_col)[machine_col].nunique().reset_index()
pivot_table = pivot_table.rename(columns={machine_col: 'Machine Count'})

# Step 2: Generate a bar graph for the pivot table
plt.figure(figsize=(12, 6))
plt.barh(pivot_table[application_col], pivot_table['Machine Count'], color='skyblue')
plt.xlabel('Number of Machines Fixed')
plt.ylabel('Applications')
plt.title('Applications vs. Machines Fixed')
plt.gca().invert_yaxis()  # Invert y-axis for better readability
plt.tight_layout()

# Save the graph to a file
graph_file = 'application_vs_machines_fixed.png'
plt.savefig(graph_file)
plt.show()

# Step 3: Use a greedy algorithm to select applications
# Target is to fix at least `target_machines` machines
application_groups = df.groupby(application_col)[machine_col].apply(list)
application_machine_count = {app: len(set(machines)) for app, machines in application_groups.items()}
sorted_applications = sorted(application_machine_count.items(), key=lambda x: x[1], reverse=True)

target_machines = 140
selected_applications = []
fixed_machines = set()

for app, count in sorted_applications:
    machines_fixed_by_app = set(application_groups[app])
    fixed_machines.update(machines_fixed_by_app)
    selected_applications.append(app)
    if len(fixed_machines) >= target_machines:
        break

# Step 4: Prepare Results DataFrame
result_df = pd.DataFrame({
    "Selected Applications": selected_applications,
    "Machines Fixed": [len(fixed_machines)] * len(selected_applications),
    "Target Machines": [target_machines] * len(selected_applications)
})

# Step 5: Save everything to a single Excel file
output_file = "patched_machines_analysis.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Original data
    df.to_excel(writer, sheet_name="Original Data", index=False)
    # Pivot table
    pivot_table.to_excel(writer, sheet_name="Pivot Table", index=False)
    # Selected applications and results
    result_df.to_excel(writer, sheet_name="Selected Applications", index=False)

print(f"\nAnalysis complete. Results saved in: {output_file}")
print(f"Graph saved as: {graph_file}")

# Print Summary
print("\nApplications to Patch to Fix Target Machines:")
print(selected_applications)
print("\nTotal Machines Fixed:", len(fixed_machines))
print(f"Target Machines: {target_machines}")
