import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def load_vulnerability_data(file_path):
    """
    Load and preprocess vulnerability dataset.
    
    Args:
        file_path (str): Path to Excel file
    
    Returns:
        pd.DataFrame: Processed vulnerability dataset
    """
    # Load dataset, drop duplicates, reset index
    df = pd.read_excel(file_path, sheet_name='QDS-above-70-crossed-40d')
    df = df.drop_duplicates(subset=['NetBIOS', 'Application'])
    
    return df

def set_cover_algorithm(df, target_machines):
    """
    Implement Set Cover Algorithm with multi-application tracking.
    
    Args:
        df (pd.DataFrame): Vulnerability dataset
        target_machines (int): Number of machines to patch
    
    Returns:
        tuple: Selected applications and their coverage details
    """
    # Create machine-to-application mapping
    machine_app_map = df.groupby('NetBIOS')['Application'].unique().to_dict()
    
    # Count how many machines each application can fix
    app_coverage = {}
    for app in df['Application'].unique():
        machines_covered = {machine for machine, apps in machine_app_map.items() if app in apps}
        app_coverage[app] = len(machines_covered)
    
    app_coverage = pd.Series(app_coverage).sort_values(ascending=False)
    
    selected_apps = []
    uncovered_machines = set(machine_app_map.keys())
    app_coverage_details = []
    
    while uncovered_machines and len(selected_apps) < len(df['Application'].unique()):
        # Select best uncovered application
        best_app = app_coverage[~app_coverage.index.isin(selected_apps)].idxmax()
        
        # Find newly covered machines by this application
        newly_covered_machines = {
            machine for machine in uncovered_machines 
            if best_app in machine_app_map[machine]
        }
        
        if newly_covered_machines:
            selected_apps.append(best_app)
            app_coverage_details.append({
                'Application': best_app,
                'Machines Fixed': len(newly_covered_machines),
                'Cumulative Machines': len(set(machine_app_map.keys()) - uncovered_machines) + len(newly_covered_machines)
            })
            
            uncovered_machines -= newly_covered_machines
        
        # Stop if machines patched meet or exceed target
        if len(machine_app_map.keys()) - len(uncovered_machines) >= target_machines:
            break
    
    return selected_apps, app_coverage_details

def create_output_visualization(app_coverage_details):
    """
    Create horizontal bar chart of application coverage.
    
    Args:
        app_coverage_details (list): Details of selected applications
    
    Returns:
        matplotlib.figure.Figure: Visualization of application coverage
    """
    plt.figure(figsize=(12, 6))
    apps = [detail['Application'] for detail in app_coverage_details]
    machines_fixed = [detail['Machines Fixed'] for detail in app_coverage_details]
    
    plt.barh(apps, machines_fixed, color='skyblue', edgecolor='navy')
    plt.xlabel('Number of Machines Fixed')
    plt.title('Application Vulnerability Coverage')
    plt.tight_layout()
    
    return plt

def main():
    # Configuration
    input_file = 'QDS-above-70-crossed-40d.xlsx'
    target_machines = 140
    output_file = 'vulnerability_analysis_output.xlsx'
    
    # Load and analyze data
    df = load_vulnerability_data(input_file)
    selected_apps, app_coverage_details = set_cover_algorithm(df, target_machines)
    
    # Create pivot table of applications and machines
    app_machine_pivot = df.groupby('Application')['NetBIOS'].nunique().reset_index()
    app_machine_pivot.columns = ['Application', 'Unique Machines']
    
    # Output to Excel
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='Original Dataset', index=False)
        app_machine_pivot.to_excel(writer, sheet_name='Application Machine Summary', index=False)
        
        coverage_df = pd.DataFrame(app_coverage_details)
        coverage_df.to_excel(writer, sheet_name='Selected Applications', index=False)
    
    # Create and save visualization
    plt = create_output_visualization(app_coverage_details)
    plt.savefig('application_coverage.png')
    plt.close()
    
    # Display summary
    print("Set Cover Analysis Results:")
    for detail in app_coverage_details:
        print(f"Application: {detail['Application']}, "
               f"Machines Fixed: {detail['Machines Fixed']}, "
               f"Cumulative Machines: {detail['Cumulative Machines']}")
    
    final_coverage = app_coverage_details[-1]['Cumulative Machines']
    print(f"\nTarget Machines: {target_machines}")
    print(f"Total Machines Fixed: {final_coverage}")
    print(f"Target Achieved: {final_coverage >= target_machines}")

if __name__ == "__main__":
    main()
