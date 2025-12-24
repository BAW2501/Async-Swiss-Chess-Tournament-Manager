import pandas as pd
import sys
import os

def clean_result(val):
    """
    Normalizes chess results.
    Converts '1/2', '0.5', 1, 0, 'Bye', '+', '-' to float scores.
    """
    val = str(val).strip()
    
    if val in ['1', '1.0', '+']:
        return 1.0
    elif val in ['0', '0.0', '-']:
        return 0.0
    elif val in ['1/2', '0.5', '½']:
        return 0.5
    elif val.lower() == 'bye':
        return 1.0 # Byes usually give 1 point
    
    # Try to convert direct numbers
    try:
        return float(val)
    except ValueError:
        return 0.0

def calculate_standings(input_file, output_file):
    print(f"Loading data from {input_file}...")
    
    # Load data (support CSV or Excel)
    if input_file.endswith('.csv'):
        df = pd.read_csv(input_file)
    else:
        df = pd.read_excel(input_file)

    # Standardize column names (strip spaces)
    df.columns = [c.strip() for c in df.columns]
    
    # Required columns check
    required_cols = ['Round', 'Board', 'White Name', 'Black Name', 'Results White', 'Results Black']
    if not all(col in df.columns for col in required_cols):
        print(f"Error: Input file must contain columns: {required_cols}")
        return

    # Data structures
    # players = { "Name": { "points": 0.0, "opponents": [], "results_vs_opponents": {} } }
    players = {}

    print("Processing games...")
    
    # 1. First Pass: Calculate Total Points and build Opponent Lists
    for index, row in df.iterrows():
        white = str(row['White Name']).strip()
        black = str(row['Black Name']).strip()
        
        # Skip empty rows
        if not white and not black:
            continue

        res_w = clean_result(row['Results White'])
        res_b = clean_result(row['Results Black'])

        # Initialize players if not exists
        if white not in players: 
            players[white] = {"points": 0.0, "opponents": [], "game_history": []}
        if black not in players and black.lower() != "bye": 
            players[black] = {"points": 0.0, "opponents": [], "game_history": []}

        # Update White
        players[white]["points"] += res_w
        if black.lower() != "bye" and black != "nan":
            players[white]["opponents"].append(black)
            players[white]["game_history"].append((black, res_w))
        
        # Update Black
        if black.lower() != "bye" and black != "nan":
            players[black]["points"] += res_b
            players[black]["opponents"].append(white)
            players[black]["game_history"].append((white, res_b))

    # 2. Second Pass: Calculate Buchholz (BucT, Buc1)
    print("Calculating Buchholz...")
    
    for p_name, p_data in players.items():
        opp_scores = []
        for opp in p_data["opponents"]:
            if opp in players:
                opp_scores.append(players[opp]["points"])
            else:
                # Handle case where opponent might not be in main list (e.g., misspelled or dropped out)
                opp_scores.append(0.0)
        
        # BucT: Total Buchholz (Sum of opponents' points)
        buc_t = sum(opp_scores)
        
        # Buc1: Buchholz Cut 1 (Sum of opponents' points - lowest score)
        if len(opp_scores) > 0:
            buc_1 = sum(sorted(opp_scores)[1:]) # Sort ascending, skip the first (lowest)
        else:
            buc_1 = 0.0
            
        p_data["BucT"] = buc_t
        p_data["Buc1"] = buc_1

    # 3. Third Pass: Calculate Direct Encounter (DE)
    # DE logic: Score calculated ONLY against players currently tied with the same points
    print("Calculating Direct Encounter...")
    
    # Group players by their score to identify ties
    score_groups = {}
    for p_name, p_data in players.items():
        pts = p_data["points"]
        if pts not in score_groups:
            score_groups[pts] = []
        score_groups[pts].append(p_name)

    for p_name, p_data in players.items():
        pts = p_data["points"]
        tied_opponents = score_groups[pts] # List of everyone with same score
        
        de_score = 0.0
        
        # Check game history against specific tied players
        for opp, result in p_data["game_history"]:
            if opp in tied_opponents:
                de_score += result
        
        p_data["DE"] = de_score

    # 4. Create Final DataFrame and Sort
    print("Finalizing Rankings...")
    
    results_list = []
    for p_name, p_data in players.items():
        results_list.append({
            "Name": p_name,
            "Points": p_data["points"],
            "DE": p_data["DE"],
            "Buc1": p_data["Buc1"],
            "BucT": p_data["BucT"]
        })

    final_df = pd.DataFrame(results_list)

    # Sorting Hierarchy: Points (Desc), DE (Desc), Buc1 (Desc), BucT (Desc)
    final_df = final_df.sort_values(
        by=['Points', 'DE', 'Buc1', 'BucT'], 
        ascending=[False, False, False, False]
    )

    # Add a Rank column
    final_df.insert(0, 'Rank', range(1, 1 + len(final_df)))

    # Save to Excel
    print(f"Saving to {output_file}...")
    final_df.to_excel(output_file, index=False)
    print("Done!")

# --- Usage ---
# Change these filenames to match your files
input_filename = 'pairings.csv' # or 'pairings.csv'
output_filename = 'standings.xlsx'

# Check if file exists before running
if os.path.exists(input_filename):
    calculate_standings(input_filename, output_filename)
else:
    # Create a dummy file for demonstration if the user runs this immediately
    print(f"File {input_filename} not found. Creating sample data...")
    data = {
        'Round': [1, 1, 2, 2],
        'Board': [1, 2, 1, 2],
        'White Name': ['Alice', 'Charlie', 'Alice', 'Bob'],
        'Black Name': ['Bob', 'David', 'Charlie', 'David'],
        'Results White': [1, 0.5, 1, 1],
        'Results Black': [0, 0.5, 0, 0]
    }
    pd.DataFrame(data).to_excel(input_filename, index=False)
    print(f"Sample created. Running script...")
    calculate_standings(input_filename, output_filename)