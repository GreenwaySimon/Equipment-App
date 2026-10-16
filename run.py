import tracker

def main():
    while True:
        print("\n=============================================")
        print("    CHAINSAW & PPE INVENTORY MANAGEMENT     ")
        print("=============================================")
        print("1. Check Out Equipment for Course")
        print("2. Return / Check In Equipment")
        print("3. Reserve Equipment for Future Assessment")
        print("4. Report Maintenance / Damage")
        print("5. Exit")
        
        choice = input("\nSelect an option (1-5): ").strip()
        
        if choice == '5':
            print("Exiting tracker. Have a great day!")
            break
            
        if choice in ['1', '2', '3', '4']:
            action_map = {
                '1': 'Check Out for Course',
                '2': 'Return / Check In',
                '3': 'Reserve for Assessment',
                '4': 'Report Maintenance / Damage'
            }
            action_type = action_map[choice]
            
            eq_ids = input("Enter Equipment IDs separated by commas (e.g. CS-01, PPE 1, TRSR - S1): ")
            trainer = input("Instructor / Assessor Name: ")
            candidate = input("Candidate / Customer Name: ")
            
            assessment_date = ""
            if choice == '3' or choice == '1':
                assessment_date = input("Assessment Date (YYYY-MM-DD or leave blank if none): ").strip()
                
            inspection = input("Inspection Result (Pass / Fail / Not Inspected): ")
            if not inspection: inspection = "Pass"
            
            notes = input("Notes / Damage Details (Optional): ")
            
            tracker.process_transaction(action_type, eq_ids, trainer, candidate, assessment_date, inspection, notes)
        else:
            print("Invalid option. Please choose between 1 and 5.")

if __name__ == '__main__':
    main()