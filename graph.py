import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def load_vulnerability_data(file_path):
    """Load vulnerability dataset, handling multi-app machines."""
    return pd.read_excel(file_path, sheet_name='QDS-above-70-crossed-40d').drop_duplicates()

def set_cover_algorithm(df, target_machines):
    """
    Implement Set Cover Algorithm with multi-application machine tracking.
    
    Args:
        df (pd.DataFrame): Vulnerability dataset
        target_machines (int): Number of machines to patch
    
    Returns:
        tuple: Selected applications and their coverage details
    """
    # Group applications by coverage
    machine_set = {row['NetBIOS']: set() for _, row in df.iterrows()}
    for _, row in df.iterrows():
        machine_set[row['NetBIOS']].add(row['Application'])
    
    app_machine_map = {}
    for app in df['Application'].unique():
        machines_covered = {machine for machine, apps in machine_set.items() if app in apps}
        app_machine_map[app] = len(machines_covered)
    
    app_machine_map = pd.Series(app_machine_map).sort_values(ascending=False)
    
    selected_apps = []
    uncovered_machines = set(machine_set.keys())
    app_coverage_details = []
    
    while uncovered_machines and len(selected_apps) < len(df['Application'].unique()):
        best_app = app_machine_map[~app_machine_map.index.isin(selected_apps)].idxmax()
        newly_covered_machines = {
            machine for machine in uncovered_machines 
            if best_app in machine_set[machine]
        }
        
        if newly_covered_machines:
            selected_apps.append(best_app)
            app_coverage_details.append({
                'Application': best_app,
                'Machines Fixed': len(newly_covered_machines),
                'Cumulative Machines': len(machine_set.keys()) - len(uncovered_machines) + len(newly_covered_machines)
            })
            
            uncovered_machines -= newly_covered_machines
        
        if len(uncovered_machines) <= target_machines:
            break
    
    return selected_apps, app_coverage_details

def main():
    # Configuration
    input_file = 'QDS-above-70-crossed-40d.xlsx'
    target_machines = 140
    output_file = 'vulnerability_analysis_output.xlsx'
    
    # Load data and run analysis
    df = load_vulnerability_data(input_file)
    selected_apps, app_coverage_details = set_cover_algorithm(df, target_machines)
    
    # Output generation remains the same as previous script
    # ... [rest of the previous script remains unchanged]

if __name__ == "__main__":
    main()
