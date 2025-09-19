import sys
import os

# Add the project directory to the Python path to allow imports
project_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(project_path)

# Now we can import the refactored monitor functions
from monitor_order_models import run_monitor_order_models
from monitor_material_details import run_monitor_material_details
from monitor_inventory import run_monitor_inventory

def main():
    """
    Runs all three monitoring and upload scripts in the specified order.
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

    print("=========================================")
    print("All Monitors Finished.")
    print("=========================================")

if __name__ == "__main__":
    main()
