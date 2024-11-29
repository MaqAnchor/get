import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def load_vulnerability_data(file_path):
    """
    Load vulnerability dataset from Excel file.
    
    Args:
        file_path (str): Path to Excel file
    
    Returns:
        pd.DataFrame: Processed vulnerability dataset
    """
    # Load specific worksheet
    df = pd.read_excel(file_path, sheet_name='QDS-above-70-crossed-40d')
    
    # Drop duplicates to ensure unique machine-application combinations
    df = df.drop_duplicates(subset=['NetBIOS', 'Application'])
    
    return df

def set_cover_algorithm(df, target_machines):
    """
    Implement Set Cover Algorithm to select minimal applications.
    
    Args:
        df (pd.DataFrame): Vulnerability dataset
        target_machines (int): Number of machines to patch
    
    Returns:
        tuple: Selected applications and their coverage details
    """
    # Group by Application and get unique machines
    app_machine_map = df.groupby('Application')['NetBIOS'].nunique().sort_values(ascending=False)
    
    selected_apps = []
    covered_machines = set()
    app_coverage_details = []
    
    while len(covered_machines) < target_machines:
        best_app = app_machine_map[~app_machine_map.index.isin(selected_apps)].idxmax()
        machines_in_app = set(df[df['Application'] == best_app]['NetBIOS'])
        
        new_machines = machines_in_app - covered_machines
        covered_machines.update(new_machines)
        
        app_coverage_details.append({
            'Application': best_app,
            'Machines Fixed': len(new_machines),
            'Cumulative Machines': len(covered_machines)
        })
        
        selected_apps.append(best_app)
        
        if len(app_machine_map) == len(selected_apps):
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
    plt.figure(figsize=(10, 6))
    apps = [detail['Application'] for detail in app_coverage_details]
    machines_fixed = [detail['Machines Fixed'] for detail in app_coverage_details]
    
    plt.barh(apps, machines_fixed)
    plt.xlabel('Number of Machines Fixed')
    plt.title('Application Vulnerability Coverage')
    plt.tight_layout()
    
    return plt

def main():
    # Configuration
    input_file = 'QDS-above-70-crossed-40d.xlsx'
    target_machines = 140
    output_file = 'vulnerability_analysis_output.xlsx'
    
    # Load data
    df = load_vulnerability_data(input_file)
    
    # Run set cover algorithm
    selected_apps, app_coverage_details = set_cover_algorithm(df, target_machines)
    
    # Create pivot table of applications and machines
    app_machine_pivot = df.groupby('Application')['NetBIOS'].nunique().reset_index()
    app_machine_pivot.columns = ['Application', 'Unique Machines']
    
    # Prepare output Excel
    with pd.ExcelWriter(output_file) as writer:
        df.to_excel(writer, sheet_name='Original Dataset', index=False)
        app_machine_pivot.to_excel(writer, sheet_name='Application Machine Summary', index=False)
        
        coverage_df = pd.DataFrame(app_coverage_details)
        coverage_df.to_excel(writer, sheet_name='Selected Applications', index=False)
    
    # Create visualization
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
