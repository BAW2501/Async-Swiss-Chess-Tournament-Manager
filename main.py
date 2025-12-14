import pandas as pd
import networkx as nx
from openpyxl import load_workbook
import random
from pathlib import Path
import sys

# Board management configuration
MAX_BOARDS = 17

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
            # Strip whitespace from column names
            pairings.columns = pairings.columns.str.strip()
            print(f"Loaded {len(pairings)} existing pairing(s)")
    except Exception as e:
        print(f"Warning: Could not load Pairings sheet ({e}), starting fresh")
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
    
    return attendees, pairings

def build_player_state(attendees, pairings):
    """Build comprehensive state for each player."""
    player_state = {}
    
    # Add all real players
    for _, row in attendees.iterrows():
        try:
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
        except Exception as e:
            print(f"Warning: Skipping invalid player row: {e}")
            continue
    
    # Debug counters
    finished_games = 0
    ongoing_games = 0
    
    # Process pairings
    for idx, pairing in pairings.iterrows():
        try:
            white_id = int(pairing["White Player"])
            black_id = int(pairing["Black Player"])
            result_white = pairing["Results White"]
            result_black = pairing["Results Black"]
            
            # Skip if players don't exist
            if white_id not in player_state or black_id not in player_state:
                print(f"Warning: Skipping pairing with unknown player(s): {white_id} vs {black_id}")
                continue
            
            # Check if game is finished
            game_finished = pd.notna(result_white) and pd.notna(result_black)
            
            if game_finished:
                finished_games += 1
            else:
                ongoing_games += 1
            
            # Mark opponents
            player_state[white_id]["opponents"].add(black_id)
            player_state[black_id]["opponents"].add(white_id)
            
            if game_finished:
                # Update points
                player_state[white_id]["points"] += float(result_white)
                player_state[black_id]["points"] += float(result_black)
                
                # Update games played
                player_state[white_id]["games_played"] += 1
                player_state[black_id]["games_played"] += 1
                
                # Update last color
                player_state[white_id]["last_color"] = "White"
                player_state[black_id]["last_color"] = "Black"
                
                # Store opponent scores for Buchholz
                player_state[white_id]["opponent_scores"].append(player_state[black_id]["points"])
                player_state[black_id]["opponent_scores"].append(player_state[white_id]["points"])
            else:
                # Game in progress
                player_state[white_id]["currently_playing"] = True
                player_state[black_id]["currently_playing"] = True
        except Exception as e:
            print(f"Warning: Error processing pairing at row {idx}: {e}")
            continue
    
    print(f"Games status: {finished_games} finished, {ongoing_games} ongoing")
    
    return player_state

def get_board_usage_count(pairings):
    """Get count of how many games are waiting/ongoing on each board."""
    board_usage = {}
    
    for _, pairing in pairings.iterrows():
        try:
            result_white = pairing["Results White"]
            result_black = pairing["Results Black"]
            game_finished = pd.notna(result_white) and pd.notna(result_black)
            
            if not game_finished:
                board_num = pairing["Board"]
                
                # Skip if board is "?" or invalid
                if pd.isna(board_num) or board_num == "?" or board_num == "":
                    continue
                
                board_num = int(board_num)
                
                if board_num not in board_usage:
                    board_usage[board_num] = 0
                board_usage[board_num] += 1
        except Exception as e:
            print(f"Warning: Error processing board usage: {e}")
            continue
    
    return board_usage

def assign_board_numbers(new_pairings, existing_pairings):
    """
    Assign board numbers with simple 2-game-per-board limit:
    1. Each board can have maximum 2 games assigned
    2. Prioritize boards with fewer games (0, then 1)
    3. Any games beyond 2-per-board capacity get "?"
    4. Next execution will reassign "?" games when boards free up
    """
    if not new_pairings:
        return []
    
    # Get board usage count (how many games currently on each board)
    board_usage = get_board_usage_count(existing_pairings)
    
    print(f"\n=== Board Assignment ===")
    print(f"Current board usage: {dict(sorted(board_usage.items()))}")
    
    # Build lists of boards by usage
    boards_with_0 = []
    boards_with_1 = []
    boards_with_2_or_more = []
    
    for board_num in range(1, MAX_BOARDS + 1):
        usage = board_usage.get(board_num, 0)
        if usage == 0:
            boards_with_0.append(board_num)
        elif usage == 1:
            boards_with_1.append(board_num)
        else:
            boards_with_2_or_more.append(board_num)
    
    print(f"Boards with 0 games: {boards_with_0}")
    print(f"Boards with 1 game: {boards_with_1}")
    print(f"Boards with 2+ games: {boards_with_2_or_more}")
    
    # Available capacity: boards with 0 or 1 game (can accept up to 2 each)
    available_capacity = len(boards_with_0) + len(boards_with_1)
    
    print(f"\nNew pairings to assign: {len(new_pairings)}")
    print(f"Available board capacity: {available_capacity} slots")
    
    # Sort new pairings by round (oldest first for priority)
    sorted_pairings = sorted(new_pairings, key=lambda x: x["Round"])
    
    board_assigned_pairings = []
    assigned_count = 0
    
    # First, use boards with 0 games
    for pairing in sorted_pairings:
        if boards_with_0:
            board_num = boards_with_0.pop(0)
            pairing["Board"] = board_num
            print(f"Round {pairing['Round']}: Assigned Board {board_num} (was empty)")
            board_usage[board_num] = 1
            assigned_count += 1
        else:
            break
    
    # Then, use boards with 1 game (filling them to 2)
    for pairing in sorted_pairings[assigned_count:]:
        if boards_with_1:
            board_num = boards_with_1.pop(0)
            pairing["Board"] = board_num
            print(f"Round {pairing['Round']}: Assigned Board {board_num} (now has 2 games)")
            board_usage[board_num] = 2
            assigned_count += 1
        else:
            break
    
    # Remaining pairings get "?" (no capacity)
    for pairing in sorted_pairings[assigned_count:]:
        pairing["Board"] = "?"
        print(f"Round {pairing['Round']}: Board ? (all boards at 2-game capacity)")
    
    for pairing in sorted_pairings:
        board_assigned_pairings.append(pairing)
    
    print(f"\nAssignment complete: {assigned_count} assigned, {len(new_pairings) - assigned_count} waiting")
    
    return board_assigned_pairings

def get_waiting_players(player_state):
    """Get list of players ready for a new pairing."""
    waiting = []
    for player_id, state in player_state.items():
        if not state["currently_playing"]:
            waiting.append(player_id)
    return waiting

def create_pairing_graph(waiting_players, player_state):
    """Create weighted graph for optimal pairing with strict color rules."""
    G = nx.Graph()
    
    # Add all waiting players as nodes
    for player_id in waiting_players:
        G.add_node(player_id)
    
    # Group players by (games_played, points)
    groups = {}
    for player_id in waiting_players:
        state = player_state[player_id]
        key = (state["games_played"], state["points"])
        if key not in groups:
            groups[key] = []
        groups[key].append(player_id)
    
    print(f"\nPlayer groups by (games, points):")
    for key, players in sorted(groups.items(), reverse=True):
        print(f"  Round {key[0]+1}, {key[1]} pts: {len(players)} player(s)")
    
    edges_added = 0
    
    # First pass: Only pair within same group
    for group_key, group_players in groups.items():
        for i, player_a in enumerate(group_players):
            for player_b in group_players[i+1:]:
                if can_pair(player_a, player_b, player_state, same_group=True):
                    weight = calculate_weight(player_a, player_b, player_state, same_group=True)
                    G.add_edge(player_a, player_b, weight=weight)
                    edges_added += 1
    
    print(f"After same-group pairing: {edges_added} edges created")
    
    # Second pass: Cross-group pairing for singletons
    singleton_players = []
    for group_key, group_players in groups.items():
        if len(group_players) == 1:
            singleton_players.append(group_players[0])
    
    if len(singleton_players) >= 2:
        print(f"Found {len(singleton_players)} singleton player(s), allowing cross-group pairing")
        for i, player_a in enumerate(singleton_players):
            for player_b in singleton_players[i+1:]:
                if can_pair(player_a, player_b, player_state, same_group=False):
                    weight = calculate_weight(player_a, player_b, player_state, same_group=False)
                    G.add_edge(player_a, player_b, weight=weight)
                    edges_added += 1
    
    print(f"Final graph: {len(waiting_players)} nodes, {edges_added} edges")
    return G

def can_pair(player_a, player_b, player_state, same_group):
    """Check if two players can be paired."""
    state_a = player_state[player_a]
    state_b = player_state[player_b]
    
    # Rule 1: No rematches
    if player_b in state_a["opponents"]:
        return False
    
    # Rule 2: Cross-group restrictions
    if not same_group:
        games_diff = abs(state_a["games_played"] - state_b["games_played"])
        if games_diff > 1:
            return False
    
    # Rule 3: Strict color alternation
    color_a = state_a["last_color"]
    color_b = state_b["last_color"]
    
    if color_a is None and color_b is None:
        return True
    elif color_a is None or color_b is None:
        return True
    elif color_a != color_b:
        return True
    
    return False

def calculate_weight(player_a, player_b, player_state, same_group):
    """Calculate edge weight for pairing priority."""
    state_a = player_state[player_a]
    state_b = player_state[player_b]
    
    if same_group:
        weight = 10000.0
    else:
        weight = 100.0
        score_diff = abs(state_a["points"] - state_b["points"])
        weight -= 50.0 * score_diff
        games_diff = abs(state_a["games_played"] - state_b["games_played"])
        weight -= 500.0 * games_diff
    
    return weight

def assign_colors(player_a, player_b, player_state):
    """Assign White and Black based on last colors."""
    color_a = player_state[player_a]["last_color"]
    color_b = player_state[player_b]["last_color"]
    
    # Both new (Round 1)
    if color_a is None and color_b is None:
        return (player_a, player_b) if random.choice([True, False]) else (player_b, player_a)
    
    # One new, one experienced
    if color_a is None:
        return (player_a, player_b) if color_b == "White" else (player_b, player_a)
    
    if color_b is None:
        return (player_b, player_a) if color_a == "White" else (player_a, player_b)
    
    # Both experienced
    return (player_b, player_a) if color_a == "White" else (player_a, player_b)

def greedy_matching(G, player_state):
    """Perform greedy maximum weight matching as a fallback."""
    edges = [(u, v, data['weight']) for u, v, data in G.edges(data=True)]
    edges.sort(key=lambda x: x[2], reverse=True)
    
    matched = set()
    matching = set()
    
    for u, v, weight in edges:
        if u not in matched and v not in matched:
            matching.add((u, v))
            matched.add(u)
            matched.add(v)
    
    return matching

def generate_pairings(player_state):
    """Generate new pairings using maximum weight matching."""
    waiting = get_waiting_players(player_state)
    
    if len(waiting) < 2:
        print(f"Only {len(waiting)} player(s) waiting. Need at least 2 for a pairing.")
        return []
    
    print(f"{len(waiting)} players waiting for pairings")
    
    # Create pairing graph
    G = create_pairing_graph(waiting, player_state)
    
    # Check if graph has edges
    if G.number_of_edges() == 0:
        print(f"⚠️  WARNING: Color clashes prevent pairing!")
        print("Players waiting:")
        for p_id in waiting:
            state = player_state[p_id]
            print(f"  - {state['name']}: Last={state['last_color']}, Pts={state['points']}")
        return []
    
    # Perform matching
    try:
        matching = nx.max_weight_matching(G, maxcardinality=True, weight='weight')
    except Exception as e:
        print(f"NetworkX matching failed ({e}), using greedy fallback")
        matching = greedy_matching(G, player_state)
    
    # Convert matching to pairings
    pairings = []
    paired_players = set()
    
    for player_a, player_b in matching:
        white_id, black_id = assign_colors(player_a, player_b, player_state)
        round_num = max(player_state[white_id]["games_played"], 
                       player_state[black_id]["games_played"]) + 1
        
        pairings.append({
            "Round": round_num,
            "White Player": white_id,
            "Black Player": black_id
        })
        
        paired_players.add(player_a)
        paired_players.add(player_b)
    
    print(f"Generated {len(pairings)} pairing(s)")
    
    # Report unpaired players
    unpaired = [p for p in waiting if p not in paired_players]
    if unpaired:
        print(f"⚠️  {len(unpaired)} player(s) remain unpaired (odd number or color conflicts)")
        for p_id in unpaired:
            state = player_state[p_id]
            print(f"  - {state['name']}")
    
    return pairings

def append_pairings_to_excel(filepath, new_pairings, player_state):
    """Append new pairings to Excel file with board assignments."""
    if not new_pairings:
        return
    
    try:
        wb = load_workbook(filepath)
        ws = wb["Pairings"]
        
        for pairing in new_pairings:
            white_id = pairing["White Player"]
            black_id = pairing["Black Player"]
            
            white_name = player_state[white_id]["name"]
            black_name = player_state[black_id]["name"]
            
            ws.append([
                pairing["Round"],
                pairing["Board"],
                white_id,
                white_name,
                black_id,
                black_name,
                None,  # Results White
                None   # Results Black
            ])
        
        wb.save(filepath)
        print(f"\n✓ Appended {len(new_pairings)} pairing(s) to Excel")
    except Exception as e:
        print(f"ERROR: Failed to save pairings: {e}")
        sys.exit(1)

def calculate_standings(player_state):
    """Calculate standings with tie-breaks."""
    standings = []
    
    for player_id, state in player_state.items():
        buchholz = sum(state["opponent_scores"])
        
        standings.append({
            "Player Name": state["name"],
            "Pt": state["points"],
            "BucT": buchholz,
            "DE": 0,
            "Ber": 0
        })
    
    standings.sort(key=lambda x: (x["Pt"], x["BucT"]), reverse=True)
    
    for i, standing in enumerate(standings):
        standing["Pos"] = i + 1
    
    return standings

def write_standings_to_excel(filepath, standings):
    """Overwrite standings sheet in Excel."""
    try:
        wb = load_workbook(filepath)
        
        if "Standings" in wb.sheetnames:
            del wb["Standings"]
        
        ws = wb.create_sheet("Standings")
        ws.append(["Pos", "Player Name", "Pt", "BucT", "DE", "Ber"])
        
        for standing in standings:
            ws.append([
                standing["Pos"],
                standing["Player Name"],
                standing["Pt"],
                standing["BucT"],
                standing["DE"],
                standing["Ber"]
            ])
        
        wb.save(filepath)
        print("✓ Updated standings")
    except Exception as e:
        print(f"ERROR: Failed to update standings: {e}")
        sys.exit(1)

def main():
    filepath = "tournament.xlsx"
    
    if not Path(filepath).exists():
        print(f"ERROR: {filepath} not found!")
        print("Please create the Excel file with 'Attendees' and 'Pairings' sheets")
        sys.exit(1)
    
    print("=" * 60)
    print("Swiss Tournament Pairing System (Production Ready)")
    print(f"Board Management: {MAX_BOARDS} boards available")
    print("=" * 60)
    
    try:
        # Load data
        attendees, pairings = load_tournament_data(filepath)
        print(f"Loaded {len(attendees)} players")
        
        # Build player state
        player_state = build_player_state(attendees, pairings)
        
        # Generate new pairings
        new_pairings = generate_pairings(player_state)
        
        # Assign board numbers
        if new_pairings:
            new_pairings = assign_board_numbers(new_pairings, pairings)
            append_pairings_to_excel(filepath, new_pairings, player_state)
        else:
            print("\nNo new pairings generated")
        
        # Calculate and write standings
        standings = calculate_standings(player_state)
        write_standings_to_excel(filepath, standings)
        
        print("\n" + "=" * 60)
        print("✓ SUCCESS! Check tournament.xlsx for updates")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n{'='*60}")
        print(f"FATAL ERROR: {e}")
        print(f"{'='*60}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()