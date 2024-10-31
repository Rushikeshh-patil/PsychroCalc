import json

# Load the JSON data from the files
with open('C:\\Git Projects\\PsychroCalc\\Generate Initial Setup List\\folder_structure_Scehdules.json') as f:
    schedules_data = json.load(f)

with open('C:\\Git Projects\\PsychroCalc\\Generate Initial Setup List\\folder_structure_Details.json') as f:
    details_data = json.load(f)
def collect_files(data, path=''):
    files = []
    name = data['name']
    if data['type'] == 'folder':
        path = path + '/' + name if path else name
        for item in data.get('contents', []):
            files.extend(collect_files(item, path))
    elif data['type'] == 'file':
        files.append({'path': path, 'name': name})
    return files

# Collect files from schedules and details
schedule_files = collect_files(schedules_data)
detail_files = collect_files(details_data)

# Define equipment types and keywords for matching
equipment_keywords = {
    # Mechanical Equipment
    'Air Handling Unit': ['AHU', 'AIR HANDLING UNIT'],
    'VAV Box': ['VAV BOX', 'VARIABLE AIR VOLUME BOX'],
    'Fan Coil Unit': ['FAN COIL UNIT', 'FCU'],
    'Exhaust Fan': ['EXHAUST FAN', 'EF'],
    'Supply Fan': ['SUPPLY FAN', 'SF'],
    'Return Fan': ['RETURN FAN', 'RF'],
    'Pump': ['PUMP'],
    'Chiller': ['CHILLER'],
    'Boiler': ['BOILER'],
    'Cooling Tower': ['COOLING TOWER', 'CT'],
    'Heat Exchanger': ['HEAT EXCHANGER', 'HX'],
    'Humidifier': ['HUMIDIFIER'],
    'VFD': ['VFD', 'VARIABLE FREQUENCY DRIVE'],
    'Expansion Tank': ['EXPANSION TANK', 'ET'],
    'Air Separator': ['AIR SEPARATOR'],
    'Unit Heater': ['UNIT HEATER', 'UH'],
    'Split System': ['SPLIT SYSTEM', 'SS'],
    'Heat Pump': ['HEAT PUMP', 'HP'],
    'Controls': ['CONTROL', 'DDC'],
    # Plumbing Equipment
    'Domestic Water System': ['DOMESTIC WATER', 'DHW', 'DCW'],
    'Sanitary Sewer System': ['SANITARY SEWER'],
    'Storm Drainage System': ['STORM DRAIN'],
    'Gas System': ['GAS METER', 'NATURAL GAS'],
    'Water Heater': ['WATER HEATER', 'WH'],
    'Backflow Preventer': ['BACKFLOW PREVENTER', 'BFP'],
    'Grease Interceptor': ['GREASE INTERCEPTOR', 'GI'],
    'Sump Pump': ['SUMP PUMP'],
    'Fixture': ['FIXTURE', 'SINK', 'TOILET', 'LAVATORY', 'URINAL'],
    'Trap Primer': ['TRAP PRIMER'],
    'Water Softener': ['WATER SOFTENER'],
    'Expansion Tank (Plumbing)': ['EXPANSION TANK'],
    'Medical Gas System': ['MEDICAL GAS', 'MED GAS'],
    'Vacuum Pump': ['VACUUM PUMP'],
    'Air Compressor': ['AIR COMPRESSOR'],
    # Add more equipment types and keywords as needed
}

# Function to map files to equipment types based on keywords
def map_files_to_equipment(files):
    mapping = {}
    for file in files:
        file_name = file['name'].lower()
        for equipment, keywords in equipment_keywords.items():
            for keyword in keywords:
                if keyword.lower() in file_name:
                    mapping.setdefault(equipment, []).append(file)
                    break  # Move to next equipment after a match
    return mapping

# Create mappings for schedules and details
equipment_to_schedules = map_files_to_equipment(schedule_files)
equipment_to_details = map_files_to_equipment(detail_files)

def ask_yes_no(question):
    while True:
        answer = input(f"{question} (Yes/No): ").strip().lower()
        if answer in ['yes', 'y']:
            return True
        elif answer in ['no', 'n']:
            return False
        else:
            print("Please enter 'Yes' or 'No'.")

def ask_choice(question, choices):
    while True:
        print(question)
        for idx, choice in enumerate(choices, 1):
            print(f"{idx}. {choice}")
        selection = input("Enter the number corresponding to your choice: ").strip()
        if selection.isdigit() and 1 <= int(selection) <= len(choices):
            return choices[int(selection) - 1]
        else:
            print("Invalid choice. Please try again.")

def ask_project_info():
    selected_equipment = set()

    # Ask about HVAC System Type
    hvac_systems = ['Single-Zone AHU', 'Multi-Zone AHU', 'Fan Coil Units', 'Split System', 'Heat Pump System']
    hvac_system = ask_choice("Select the type of HVAC system:", hvac_systems)
    
    if hvac_system == 'Single-Zone AHU':
        selected_equipment.add('Air Handling Unit')
        if ask_yes_no("Does the AHU have a VFD?"):
            selected_equipment.add('VFD')
    elif hvac_system == 'Multi-Zone AHU':
        selected_equipment.add('Air Handling Unit')
        selected_equipment.add('VAV Box')
        if ask_yes_no("Does the AHU have a VFD?"):
            selected_equipment.add('VFD')
    elif hvac_system == 'Fan Coil Units':
        selected_equipment.add('Fan Coil Unit')
    elif hvac_system == 'Split System':
        selected_equipment.add('Split System')
    elif hvac_system == 'Heat Pump System':
        selected_equipment.add('Heat Pump')

    # Common equipment
    if ask_yes_no("Does the project include Exhaust Fans?"):
        selected_equipment.add('Exhaust Fan')
    if ask_yes_no("Does the project include Controls?"):
        selected_equipment.add('Controls')

    # Heating System
    if ask_yes_no("Does the project include a Heating System?"):
        heating_types = ['Boiler', 'Heat Exchanger']
        heating_choice = ask_choice("Select the type of Heating System:", heating_types)
        selected_equipment.add(heating_choice)
        selected_equipment.add('Pump')
        selected_equipment.add('Expansion Tank')
        selected_equipment.add('Air Separator')
        if ask_yes_no("Are there Unit Heaters in the project?"):
            selected_equipment.add('Unit Heater')

    # Cooling System
    if ask_yes_no("Does the project include a Cooling System?"):
        cooling_types = ['Chiller', 'Heat Exchanger']
        cooling_choice = ask_choice("Select the type of Cooling System:", cooling_types)
        selected_equipment.add(cooling_choice)
        selected_equipment.add('Pump')
        selected_equipment.add('Expansion Tank')
        selected_equipment.add('Air Separator')
        if cooling_choice == 'Chiller':
            if ask_yes_no("Does the project include a Cooling Tower?"):
                selected_equipment.add('Cooling Tower')

    # Plumbing Systems
    if ask_yes_no("Does the project include Plumbing Systems?"):
        # Domestic Water System
        if ask_yes_no("Include Domestic Water System?"):
            selected_equipment.add('Domestic Water System')
            selected_equipment.add('Backflow Preventer')
            if ask_yes_no("Does the project include Water Heaters?"):
                wh_types = ['Gas Water Heater', 'Electric Water Heater']
                wh_choice = ask_choice("Select the type of Water Heater:", wh_types)
                selected_equipment.add(wh_choice)
            if ask_yes_no("Does the project include Water Softeners?"):
                selected_equipment.add('Water Softener')
            if ask_yes_no("Does the project include Fixtures (sinks, toilets, etc.)?"):
                selected_equipment.add('Fixture')
            if ask_yes_no("Does the project include Trap Primers?"):
                selected_equipment.add('Trap Primer')
            if ask_yes_no("Does the project include Expansion Tanks?"):
                selected_equipment.add('Expansion Tank (Plumbing)')
        # Sanitary Sewer System
        if ask_yes_no("Include Sanitary Sewer System?"):
            selected_equipment.add('Sanitary Sewer System')
            if ask_yes_no("Does the project include Grease Interceptors?"):
                selected_equipment.add('Grease Interceptor')
            if ask_yes_no("Does the project include Sump Pumps?"):
                selected_equipment.add('Sump Pump')
        # Storm Drainage System
        if ask_yes_no("Include Storm Drainage System?"):
            selected_equipment.add('Storm Drainage System')
        # Gas System
        if ask_yes_no("Include Gas System?"):
            selected_equipment.add('Gas System')
        # Medical Gas System
        if ask_yes_no("Include Medical Gas System?"):
            selected_equipment.add('Medical Gas System')
            if ask_yes_no("Does the project include Vacuum Pumps?"):
                selected_equipment.add('Vacuum Pump')
            if ask_yes_no("Does the project include Air Compressors?"):
                selected_equipment.add('Air Compressor')

    return selected_equipment

def generate_required_lists(selected_equipment):
    required_schedules = []
    required_details = []

    for equipment in selected_equipment:
        # Get schedules
        schedules = equipment_to_schedules.get(equipment)
        if schedules:
            required_schedules.extend(schedules)
        # Get details
        details = equipment_to_details.get(equipment)
        if details:
            required_details.extend(details)

    return required_schedules, required_details

def main():
    selected_equipment = ask_project_info()
    schedules, details = generate_required_lists(selected_equipment)

    print("\nRequired Schedules:")
    for schedule in sorted(set([s['name'] for s in schedules])):  # Using set to avoid duplicates
        print(f"- {schedule}")

    print("\nRequired Details:")
    for detail in sorted(set([d['name'] for d in details])):
        print(f"- {detail}")

if __name__ == "__main__":
    main()