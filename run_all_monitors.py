import sys
import os

# Add the project directory to the Python path to allow imports
project_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(project_path)

# Now we can import the refactored monitor functions
from monitor_order_models import run_monitor_order_models
from monitor_material_details import run_monitor_material_details
from monitor_inventory import run_monitor_inventory
from monitor_semi_finished import run_monitor_semi_finished
from monitor_semi_finished_model_db import run_monitor_semi_finished_model_db
from sync_estimated_arrival_dates import run_sync_estimated_arrival_dates

def main():
    """
    Runs all five monitoring and upload scripts in the specified order.
    """
    print("=========================================")
    print("Starting All Automatic Upload Monitors...")
    print("=========================================\n")

    # 1. Order & Model
    try:
        run_monitor_order_models()
    except Exception as e:
        print("\nAn error occurred during the Order & Model upload:")
        print(f"Error: {e}\n")

    print("\n-----------------------------------------\n")

    # 2. Material Details
    try:
        run_monitor_material_details()
    except Exception as e:
        print("\nAn error occurred during the Material Details upload:")
        print(f"Error: {e}\n")

    print("\n-----------------------------------------\n")

    # 3. Inventory
    try:
        run_monitor_inventory()
    except Exception as e:
        print("\nAn error occurred during the Inventory upload:")
        print(f"Error: {e}\n")

    print("\n-----------------------------------------\n")

    # 4. Semi-Finished
    try:
        run_monitor_semi_finished()
    except Exception as e:
        print("\nAn error occurred during the Semi-Finished upload:")
        print(f"Error: {e}\n")

    print("\n-----------------------------------------\n")

    # 5. Semi-Finished Model DB
    try:
        run_monitor_semi_finished_model_db()
    except Exception as e:
        print("\nAn error occurred during the Semi-Finished Model DB upload:")
        print(f"Error: {e}\n")

    print("\n-----------------------------------------\n")

    # 6. Sync Estimated Arrival Dates
    try:
        run_sync_estimated_arrival_dates()
    except Exception as e:
        print("\nAn error occurred during the Sync Estimated Arrival Dates:")
        print(f"Error: {e}\n")

    print("=========================================")
    print("All Monitors Finished.")
    print("=========================================")

if __name__ == "__main__":
    main()
