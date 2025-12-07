import pandas as pd
import networkx as nx
from openpyxl import load_workbook
import random
from pathlib import Path
# TODO work with limited amount of boards (board management) prio for current round later rounds
# TODO some sort of html export with beautiful UI
def load_tournament_data(filepath="tournament.xlsx"):
    """Load attendees and pairings from Excel file."""
    attendees = pd.read_excel(filepath, sheet_name="Attendees")
    try:
        pairings = pd.read_excel(filepath, sheet_name="Pairings")
        if pairings.empty:
            pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
    except:
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
    
    return attendees, pairings

def build_player_state(attendees, pairings):
    """Build comprehensive state for each player."""
    player_state = {}
    
    for _, row in attendees.iterrows():
        player_id = row["ID"]
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
        white_id = pairing["White Player"]
        black_id = pairing["Black Player"]
        result_white = pairing["Results White"]
        result_black = pairing["Results Black"]
        
        # Check if game is finished
        game_finished = pd.notna(result_white) and pd.notna(result_black)
        
        if white_id in player_state and black_id in player_state:
            # Mark opponents
            player_state[white_id]["opponents"].add(black_id)
            player_state[black_id]["opponents"].add(white_id)
            
            if game_finished:
                # Update points
                player_state[white_id]["points"] += result_white
                player_state[black_id]["points"] += result_black
                
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
    
    return player_state

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
    
    print(f"Player groups by (games, points):")
    for key, players in sorted(groups.items(), reverse=True):
        print(f"  Round {key[0]+1}, {key[1]} pts: {len(players)} player(s)")
    
    edges_added = 0
    
    # First pass: Only pair within same group (same games played AND same points)
    for group_key, group_players in groups.items():
        for i, player_a in enumerate(group_players):
            for player_b in group_players[i+1:]:
                if can_pair(player_a, player_b, player_state, same_group=True):
                    weight = calculate_weight(player_a, player_b, player_state, same_group=True)
                    G.add_edge(player_a, player_b, weight=weight)
                    edges_added += 1
    
    print(f"After same-group pairing: {edges_added} edges created.")
    
    # Second pass: If we have groups with only 1 player, allow cross-group pairing
    # But ONLY as a last resort
    singleton_players = []
    for group_key, group_players in groups.items():
        if len(group_players) == 1:
            singleton_players.append(group_players[0])
    
    if len(singleton_players) >= 2:
        print(f"Found {len(singleton_players)} singleton player(s), allowing cross-group pairing as last resort.")
        for i, player_a in enumerate(singleton_players):
            for player_b in singleton_players[i+1:]:
                if can_pair(player_a, player_b, player_state, same_group=False):
                    # Much lower weight for cross-group pairings
                    weight = calculate_weight(player_a, player_b, player_state, same_group=False)
                    G.add_edge(player_a, player_b, weight=weight)
                    edges_added += 1
    
    print(f"Final graph: {len(waiting_players)} nodes, {edges_added} edges.")
    return G

def can_pair(player_a, player_b, player_state, same_group):
    """Check if two players can be paired."""
    state_a = player_state[player_a]
    state_b = player_state[player_b]
    
    # Rule 1: No rematches
    if player_b in state_a["opponents"]:
        return False
    
    # Rule 2: If not same group, only pair if truly necessary
    if not same_group:
        # Only allow cross-group pairing if games played differ by at most 1
        games_diff = abs(state_a["games_played"] - state_b["games_played"])
        if games_diff > 1:
            return False
    
    # Rule 3: Strict color alternation
    color_a = state_a["last_color"]
    color_b = state_b["last_color"]
    
    # Check if colors are compatible
    if color_a is None and color_b is None:
        # Both new players (Round 1)
        return True
    elif color_a is None or color_b is None:
        # One new, one experienced - always compatible
        return True
    elif color_a != color_b:
        # Both experienced with opposite last colors
        return True
    
    return False

def calculate_weight(player_a, player_b, player_state, same_group):
    """Calculate edge weight for pairing priority."""
    state_a = player_state[player_a]
    state_b = player_state[player_b]
    
    if same_group:
        # High base weight for same-group pairings
        weight = 10000.0
    else:
        # Low base weight for cross-group pairings (last resort)
        weight = 100.0
        
        # Additional penalties for cross-group
        score_diff = abs(state_a["points"] - state_b["points"])
        weight -= 50.0 * score_diff
        
        games_diff = abs(state_a["games_played"] - state_b["games_played"])
        weight -= 500.0 * games_diff
    
    return weight

def assign_colors(player_a, player_b, player_state):
    """Assign White and Black based on last colors."""
    color_a = player_state[player_a]["last_color"]
    color_b = player_state[player_b]["last_color"]
    
    # Both new (Round 1) - assign randomly
    if color_a is None and color_b is None:
        if random.choice([True, False]):
            return player_a, player_b
        else:
            return player_b, player_a
    
    # One new, one experienced
    if color_a is None:
        # B played last, so assign opposite
        if color_b == "White":
            return player_a, player_b  # A=White, B=Black
        else:
            return player_b, player_a  # B=White, A=Black
    
    if color_b is None:
        # A played last, so assign opposite
        if color_a == "White":
            return player_b, player_a  # B=White, A=Black
        else:
            return player_a, player_b  # A=White, B=Black
    
    # Both experienced - assign opposite of last color
    if color_a == "White":
        return player_b, player_a  # B=White, A=Black
    else:
        return player_a, player_b  # A=White, B=Black

def greedy_matching(G, player_state):
    """Perform greedy maximum weight matching as a fallback."""
    # Sort edges by weight (descending)
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
    
    print(f"{len(waiting)} players waiting for pairings...")
    
    # Create pairing graph
    G = create_pairing_graph(waiting, player_state)
    
    # Check if graph has any edges
    if G.number_of_edges() == 0:
        print(f"⚠️  WARNING: {len(waiting)} players waiting, but color clashes prevent pairing!")
        print("Players waiting:")
        for p_id in waiting:
            state = player_state[p_id]
            print(f"  - {state['name']}: Last color = {state['last_color']}, Points = {state['points']}")
        return []
    
    # Try NetworkX matching first, with fallback to greedy
    try:
        matching = nx.max_weight_matching(G, maxcardinality=True, weight='weight')
    except (AssertionError, Exception) as e:
        print(f"NetworkX matching failed ({e}), using greedy fallback...")
        matching = greedy_matching(G, player_state)
    
    if len(matching) == 0 and len(waiting) >= 2:
        print(f"⚠️  WARNING: {len(waiting)} players waiting, but color clashes prevent pairing!")
        print("Players waiting:")
        for p_id in waiting:
            state = player_state[p_id]
            print(f"  - {state['name']}: Last color = {state['last_color']}, Points = {state['points']}")
        return []
    
    # Convert matching to pairings with color assignment
    pairings = []
    for player_a, player_b in matching:
        white_id, black_id = assign_colors(player_a, player_b, player_state)
        
        # Determine round number (max of both players + 1)
        round_num = max(player_state[white_id]["games_played"], 
                       player_state[black_id]["games_played"]) + 1
        
        pairings.append({
            "Round": round_num,
            "White Player": white_id,
            "Black Player": black_id
        })
    
    print(f"Generated {len(pairings)} new pairing(s).")
    
    # Report unpaired players if any
    paired_players = set()
    for pairing in pairings:
        paired_players.add(pairing["White Player"])
        paired_players.add(pairing["Black Player"])
    
    unpaired = [p for p in waiting if p not in paired_players]
    if unpaired:
        print(f"{len(unpaired)} player(s) remain unpaired (odd number or color conflicts).")
    
    return pairings

def append_pairings_to_excel(filepath, new_pairings, player_state):
    """Append new pairings to Excel file."""
    if not new_pairings:
        return
    
    wb = load_workbook(filepath)
    ws = wb["Pairings"]
    
    # Find next board number
    max_board = 0
    for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if row[1] is not None:
            max_board = max(max_board, row[1])
    
    # Append new pairings
    for i, pairing in enumerate(new_pairings):
        board_num = max_board + i + 1
        white_id = pairing["White Player"]
        black_id = pairing["Black Player"]
        
        ws.append([
            pairing["Round"],
            board_num,
            white_id,
            player_state[white_id]["name"],  # White Name
            black_id,
            player_state[black_id]["name"],  # Black Name
            None,  # Results White
            None   # Results Black
        ])
    
    wb.save(filepath)
    print(f"Appended {len(new_pairings)} pairing(s) to Excel.")

def calculate_standings(player_state):
    """Calculate standings with tie-breaks."""
    standings = []
    
    for player_id, state in player_state.items():
        # Calculate Buchholz (sum of opponents' scores)
        buchholz = sum(state["opponent_scores"])
        
        standings.append({
            "Player Name": state["name"],
            "Pt": state["points"],
            "BucT": buchholz,
            "DE": 0,  # Direct Encounter (placeholder)
            "Ber": 0  # Berger (placeholder)
        })
    
    # Sort by points (desc), then Buchholz (desc)
    standings.sort(key=lambda x: (x["Pt"], x["BucT"]), reverse=True)
    
    # Add position
    for i, standing in enumerate(standings):
        standing["Pos"] = i + 1
    
    return standings

def write_standings_to_excel(filepath, standings):
    """Overwrite standings sheet in Excel."""
    wb = load_workbook(filepath)
    
    # Remove old standings sheet if exists
    if "Standings" in wb.sheetnames:
        del wb["Standings"]
    
    # Create new standings sheet
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
    print("Updated standings.")

def main():
    filepath = "tournament.xlsx"
    
    if not Path(filepath).exists():
        print(f"Error: {filepath} not found!")
        print("Please create the Excel file with 'Attendees' and 'Pairings' sheets.")
        return
    
    print("=" * 60)
    print("Swiss Tournament Pairing System")
    print("=" * 60)
    
    # Load data
    attendees, pairings = load_tournament_data(filepath)
    print(f"Loaded {len(attendees)} players.")
    
    # Build player state
    player_state = build_player_state(attendees, pairings)
    
    # Generate new pairings
    new_pairings = generate_pairings(player_state)
    
    # Write to Excel
    if new_pairings:
        append_pairings_to_excel(filepath, new_pairings, player_state)
    
    # Calculate and write standings
    standings = calculate_standings(player_state)
    write_standings_to_excel(filepath, standings)
    
    print("=" * 60)
    print("Done! Check tournament.xlsx for updated pairings and standings.")
    print("=" * 60)

if __name__ == "__main__":
    main()