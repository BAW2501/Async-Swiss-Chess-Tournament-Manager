import pandas as pd
from openpyxl import load_workbook
import random
from pathlib import Path
import sys

MAX_BOARDS = 30

def load_tournament_data(filepath="tournament.xlsx"):
    """Load attendees and pairings from Excel file."""
    try:
        attendees = pd.read_excel(filepath, sheet_name="Attendees")
    except Exception as e:
        print(f"ERROR: Failed to load Attendees sheet: {e}")
        sys.exit(1)
    
    try:
        pairings = pd.read_excel(filepath, sheet_name="Pairings")
        if pairings.empty:
            pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        else:
            pairings.columns = pairings.columns.str.strip()
            print(f"Loaded {len(pairings)} existing pairing(s)")
    except Exception as e:
        print(f"Warning: Could not load Pairings sheet ({e}), starting fresh")
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
    
    return attendees, pairings

def build_player_state(attendees, pairings):
    """Build state for each player."""
    player_state = {}
    
    # Initialize all players
    for _, row in attendees.iterrows():
        player_id = int(row["ID"])
        player_state[player_id] = {
            "name": f"{row['First Name']} {row['Last Name']}",
            "points": 0.0,
            "games_played": 0,
            "last_color": None,
            "opponents": set(),
            "opponent_scores": [],
            "currently_playing": False
        }
    
    # Process pairings
    for _, pairing in pairings.iterrows():
        white_id = int(pairing["White Player"])
        black_id = int(pairing["Black Player"])
        result_white = pairing["Results White"]
        result_black = pairing["Results Black"]
        
        if white_id not in player_state or black_id not in player_state:
            continue
        
        game_finished = pd.notna(result_white) and pd.notna(result_black)
        
        player_state[white_id]["opponents"].add(black_id)
        player_state[black_id]["opponents"].add(white_id)
        
        if game_finished:
            player_state[white_id]["points"] += float(result_white)
            player_state[black_id]["points"] += float(result_black)
            player_state[white_id]["games_played"] += 1
            player_state[black_id]["games_played"] += 1
            player_state[white_id]["last_color"] = "White"
            player_state[black_id]["last_color"] = "Black"
            player_state[white_id]["opponent_scores"].append(black_id)
            player_state[black_id]["opponent_scores"].append(white_id)
            player_state[white_id]["currently_playing"] = False
            player_state[black_id]["currently_playing"] = False
        else:
            player_state[white_id]["currently_playing"] = True
            player_state[black_id]["currently_playing"] = True
    
    # Calculate Buchholz scores
    for player_id in player_state:
        opponent_ids = player_state[player_id]["opponent_scores"]
        player_state[player_id]["opponent_scores"] = [
            player_state[opp_id]["points"] for opp_id in opponent_ids
        ]
    
    return player_state

def get_board_usage_count(pairings):
    """Get count of ongoing games per board."""
    board_usage = {}
    
    for _, pairing in pairings.iterrows():
        result_white = pairing["Results White"]
        result_black = pairing["Results Black"]
        game_finished = pd.notna(result_white) and pd.notna(result_black)
        
        if not game_finished:
            board_num = pairing["Board"]
            if pd.notna(board_num) and board_num != "?":
                board_num = int(board_num)
                board_usage[board_num] = board_usage.get(board_num, 0) + 1
    
    return board_usage

def assign_board_numbers(new_pairings, existing_pairings):
    """Assign board numbers using round-robin."""
    if not new_pairings:
        return []
    
    board_usage = get_board_usage_count(existing_pairings)
    current_board = 0
    
    for pairing in new_pairings:
        assigned = False
        
        for _ in range(MAX_BOARDS):
            current_board = (current_board % MAX_BOARDS) + 1
            
            if board_usage.get(current_board, 0) < 2:
                pairing["Board"] = current_board
                board_usage[current_board] = board_usage.get(current_board, 0) + 1
                assigned = True
                break
        
        if not assigned:
            pairing["Board"] = "?"
    
    return new_pairings

def can_pair(p1_state, p2_state, p2_id):
    """Check if two players can be paired."""
    # No rematches
    if p2_id in p1_state["opponents"]:
        return False
    # must have played same number of games
    if p1_state["games_played"] != p2_state["games_played"]:
        return False
    
    # Color alternation check
    color1 = p1_state["last_color"]
    color2 = p2_state["last_color"]
    
    # If either player is new, they can pair
    if color1 is None or color2 is None:
        return True
    

    # Both must have different last colors
    return color1 != color2

def assign_colors(p1_id, p2_id, player_state):
    """Assign colors based on last color played."""
    color1 = player_state[p1_id]["last_color"]
    color2 = player_state[p2_id]["last_color"]
    
    if color1 is None and color2 is None:
        return (p1_id, p2_id) if random.choice([True, False]) else (p2_id, p1_id)
    
    if color1 is None:
        return (p1_id, p2_id) if color2 == "White" else (p2_id, p1_id)
    
    if color2 is None:
        return (p2_id, p1_id) if color1 == "White" else (p1_id, p2_id)
    
    # Both have history: player who had White last plays Black
    return (p2_id, p1_id) if color1 == "White" else (p1_id, p2_id)

def generate_pairings(player_state):
    """Generate pairings using simple greedy approach."""
    # Get waiting players
    waiting = [pid for pid, state in player_state.items() if not state["currently_playing"]]
    
    if len(waiting) < 2:
        print(f"\nOnly {len(waiting)} player(s) ready for pairing.")
        return []
    
    print(f"\n{'='*60}")
    print(f"PAIRING: {len(waiting)} players waiting")
    print(f"{'='*60}")
    
    # Sort by: games_played (ascending), then points (descending)
    waiting.sort(key=lambda p: (player_state[p]["games_played"], -player_state[p]["points"]))
    
    pairings = []
    paired = set()
    
    # Simple greedy matching
    for i, p1 in enumerate(waiting):
        if p1 in paired:
            continue
        
        p1_state = player_state[p1]
        
        # Try to find a suitable opponent
        for p2 in waiting[i+1:]:
            if p2 in paired:
                continue
            
            p2_state = player_state[p2]
            
            # Check if they can be paired
            if can_pair(p1_state, p2_state, p2):
                # Pair them
                white_id, black_id = assign_colors(p1, p2, player_state)
                round_num = max(p1_state["games_played"], p2_state["games_played"]) + 1
                
                pairings.append({
                    "Round": round_num,
                    "White Player": white_id,
                    "Black Player": black_id
                })
                
                paired.add(p1)
                paired.add(p2)
                break
    
    unpaired = len(waiting) - len(paired)
    print(f"Generated {len(pairings)} pairing(s), {unpaired} player(s) unpaired")
    
    return pairings

def append_pairings_to_excel(filepath, new_pairings, player_state):
    """Append new pairings to Excel."""
    if not new_pairings:
        return
    
    wb = load_workbook(filepath)
    ws = wb["Pairings"]
    
    for pairing in new_pairings:
        white_id = pairing["White Player"]
        black_id = pairing["Black Player"]
        
        ws.append([
            pairing["Round"],
            pairing["Board"],
            white_id,
            player_state[white_id]["name"],
            black_id,
            player_state[black_id]["name"],
            None,
            None
        ])
    
    wb.save(filepath)
    print(f"✓ Appended {len(new_pairings)} pairing(s)")

def calculate_standings(player_state):
    """Calculate standings with Buchholz and Berger tie-breaks."""
    standings = []
    
    for player_id, state in player_state.items():
        buchholz = sum(state["opponent_scores"])
        
        # Berger: calculate from opponent_scores list
        berger = 0.0
        # This is a simplified Berger - you'd need game results for exact calculation
        
        standings.append({
            "Pos": 0,
            "Player Name": state["name"],
            "Pt": state["points"],
            "BucT": buchholz,
            "Ber": berger
        })
    
    standings.sort(key=lambda x: (x["Pt"], x["BucT"], x["Ber"]), reverse=True)
    
    for i, standing in enumerate(standings):
        standing["Pos"] = i + 1
    
    return standings

def write_standings_to_excel(filepath, standings):
    """Write standings to Excel."""
    wb = load_workbook(filepath)
    
    if "Standings" in wb.sheetnames:
        del wb["Standings"]
    
    ws = wb.create_sheet("Standings")
    ws.append(["Pos", "Player Name", "Pt", "BucT", "Ber"])
    
    for standing in standings:
        ws.append([
            standing["Pos"],
            standing["Player Name"],
            standing["Pt"],
            standing["BucT"],
            standing["Ber"]
        ])
    
    wb.save(filepath)
    print("✓ Updated standings")

def main():
    filepath = "tournament.xlsx"
    
    if not Path(filepath).exists():
        print(f"ERROR: {filepath} not found!")
        sys.exit(1)
    
    print("=" * 60)
    print("Swiss Tournament Pairing System (Simplified)")
    print("=" * 60)
    
    attendees, pairings = load_tournament_data(filepath)
    player_state = build_player_state(attendees, pairings)
    
    new_pairings = generate_pairings(player_state)
    
    if new_pairings:
        new_pairings = assign_board_numbers(new_pairings, pairings)
        append_pairings_to_excel(filepath, new_pairings, player_state)
    
    standings = calculate_standings(player_state)
    write_standings_to_excel(filepath, standings)
    
    print("\n" + "=" * 60)
    print("✓ SUCCESS!")
    print("=" * 60)

if __name__ == "__main__":
    main()