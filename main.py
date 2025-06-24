from utils.excel_handler import create_condition_scores_sheet

def run():
    outputs = create_condition_scores_sheet(
        output_path="reports",
        assets_path="data/assets.xlsx",
        incidents_path="data/incidents.xlsx",
        tasks_path="data/tasks.xlsx",
        config_path="condition_scoring.json",
        export_mapping=True
    )
    print("Report:", outputs['report'])
    print("GeoJSON:", outputs['mapping']['geojson'])
    print("Mapping Excel:", outputs['mapping']['excel'])

if __name__ == "__main__":
    run()