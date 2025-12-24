import unittest
import pandas as pd
from unittest.mock import Mock, patch, MagicMock


# === ASSUMPTION ===
# Assuming the original code is saved in a file named 'main.py'
# If your file is named something else, change this import.
from main import (
    build_player_state, 
    get_board_usage_count, 
    assign_board_numbers,
    can_pair,
    assign_colors,
    generate_pairings,
    calculate_standings,
    MAX_BOARDS
)

class TestBuildPlayerState(unittest.TestCase):
    """Test player state building from attendees and pairings."""
    
    def setUp(self):
        """Create sample data for testing."""
        self.attendees = pd.DataFrame({
            'ID': [1, 2, 3, 4],
            'First Name': ['Alice', 'Bob', 'Charlie', 'David'],
            'Last Name': ['Smith', 'Jones', 'Brown', 'Wilson']
        })
    
    def test_empty_pairings(self):
        """Test building state with no games played."""
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        player_state = build_player_state(self.attendees, pairings)
        
        # Check initialization
        self.assertEqual(player_state[1]["points"], 0.0)
        self.assertEqual(player_state[1]["games_played"], 0)
        self.assertFalse(player_state[1]["currently_playing"])
        self.assertIsNone(player_state[1]["last_color"])
        self.assertEqual(len(player_state), 4)

    def test_single_finished_game(self):
        """Test state after one completed game."""
        pairings = pd.DataFrame({
            'Round': [1],
            'Board': [1],
            'White Player': [1],
            'Black Player': [2],
            'Results White': [1.0],
            'Results Black': [0.0]
        })
        
        player_state = build_player_state(self.attendees, pairings)
        
        # Player 1 (Winner)
        self.assertEqual(player_state[1]["points"], 1.0)
        self.assertEqual(player_state[1]["games_played"], 1)
        self.assertEqual(player_state[1]["last_color"], "White")
        self.assertIn(2, player_state[1]["opponents"])
        
        # Player 2 (Loser)
        self.assertEqual(player_state[2]["points"], 0.0)
        self.assertEqual(player_state[2]["last_color"], "Black")
        self.assertIn(1, player_state[2]["opponents"])
    
    def test_ongoing_game(self):
        """Test state with game in progress."""
        pairings = pd.DataFrame({
            'Round': [1],
            'Board': [1],
            'White Player': [1],
            'Black Player': [2],
            'Results White': [None],
            'Results Black': [None]
        })
        
        player_state = build_player_state(self.attendees, pairings)
        
        self.assertTrue(player_state[1]["currently_playing"])
        self.assertTrue(player_state[2]["currently_playing"])
        self.assertEqual(player_state[1]["points"], 0.0)
        self.assertEqual(player_state[1]["games_played"], 0)
        # Opponents are added even if game isn't finished in this logic
        self.assertIn(2, player_state[1]["opponents"]) 
    
    def test_mixed_finished_and_ongoing(self):
        """Test with both finished and ongoing games."""
        pairings = pd.DataFrame({
            'Round': [1, 1, 2],
            'Board': [1, 2, 3],
            'White Player': [1, 3, 1],
            'Black Player': [2, 4, 3],
            'Results White': [1.0, 0.5, None],
            'Results Black': [0.0, 0.5, None]
        })
        
        player_state = build_player_state(self.attendees, pairings)
        
        # Player 1: Won R1, Playing R2
        self.assertEqual(player_state[1]["points"], 1.0)
        self.assertTrue(player_state[1]["currently_playing"])
        
        # Player 2: Lost R1, Waiting
        self.assertEqual(player_state[2]["points"], 0.0)
        self.assertFalse(player_state[2]["currently_playing"])
        
        # Player 3: Drew R1, Playing R2
        self.assertEqual(player_state[3]["points"], 0.5)
        self.assertTrue(player_state[3]["currently_playing"])
    
    def test_buchholz_calculation(self):
        """Test that opponent scores are correctly stored for Buchholz."""
        # P1 beats P3 (P3 has 0 pts)
        # P2 draws P4 (P4 has 0.5 pts)
        pairings = pd.DataFrame({
            'Round': [1, 1],
            'Board': [1, 2],
            'White Player': [1, 2],
            'Black Player': [3, 4],
            'Results White': [1.0, 0.5],
            'Results Black': [0.0, 0.5]
        })
        
        player_state = build_player_state(self.attendees, pairings)
        
        # Player 1 played Player 3. Player 3 has 0.0 points.
        self.assertEqual(player_state[1]["opponent_scores"], [0.0])
        
        # Player 3 played Player 1. Player 1 has 1.0 points.
        self.assertEqual(player_state[3]["opponent_scores"], [1.0])
        
        # Player 2 played Player 4. Player 4 has 0.5 points.
        self.assertEqual(player_state[2]["opponent_scores"], [0.5])

    def test_invalid_player_ids(self):
        """Test handling of invalid player IDs in pairings."""
        pairings = pd.DataFrame({
            'Round': [1],
            'Board': [1],
            'White Player': [1],
            'Black Player': [999],  # Player doesn't exist
            'Results White': [1.0],
            'Results Black': [0.0]
        })
        
        # Should not crash
        player_state = build_player_state(self.attendees, pairings)
        
        # Player 1 should exist but game shouldn't be recorded effectively because logic requires both IDs
        # The logic: if white_id not in player_state or black_id not in player_state: continue
        self.assertEqual(player_state[1]["games_played"], 0)
        self.assertEqual(player_state[1]["points"], 0.0)

    def test_draw_results(self):
        """Test handling of draw (0.5-0.5) results."""
        pairings = pd.DataFrame({
            'Round': [1],
            'Board': [1],
            'White Player': [1],
            'Black Player': [2],
            'Results White': [0.5],
            'Results Black': [0.5]
        })
        
        player_state = build_player_state(self.attendees, pairings)
        self.assertEqual(player_state[1]["points"], 0.5)
        self.assertEqual(player_state[2]["points"], 0.5)


class TestGetBoardUsageCount(unittest.TestCase):
    """Test board usage counting."""
    
    def test_empty_pairings(self):
        pairings = pd.DataFrame(columns=["Round", "Board", "Results White", "Results Black"])
        usage = get_board_usage_count(pairings)
        self.assertEqual(usage, {})
    
    def test_all_finished_games(self):
        pairings = pd.DataFrame({
            'Board': [1, 2],
            'Results White': [1.0, 0.5],
            'Results Black': [0.0, 0.5]
        })
        usage = get_board_usage_count(pairings)
        self.assertEqual(usage, {})
    
    def test_ongoing_games_single_board(self):
        pairings = pd.DataFrame({
            'Board': [1, 1],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        usage = get_board_usage_count(pairings)
        self.assertEqual(usage, {1: 2})
    
    def test_question_mark_boards(self):
        pairings = pd.DataFrame({
            'Board': ['?', 1],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        usage = get_board_usage_count(pairings)
        self.assertEqual(usage, {1: 1})


class TestAssignBoardNumbers(unittest.TestCase):
    """Test board number assignment."""
    
    def setUp(self):
        self.empty_existing = pd.DataFrame(columns=["Results White", "Results Black", "Board"])

    def test_empty_pairings(self):
        new_pairings = []
        result = assign_board_numbers(new_pairings, self.empty_existing)
        self.assertEqual(result, [])
    
    def test_assign_to_empty_boards(self):
        new_pairings = [{"White Player": 1}, {"White Player": 2}]
        result = assign_board_numbers(new_pairings, self.empty_existing)
        
        self.assertEqual(result[0]["Board"], 1)
        self.assertEqual(result[1]["Board"], 2)
    
    def test_skip_busy_boards(self):
        # Board 1 has 2 ongoing games
        existing = pd.DataFrame({
            'Board': [1, 1],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        
        new_pairings = [{"White Player": 5}]
        result = assign_board_numbers(new_pairings, existing)
        
        # Should skip 1 and go to 2
        self.assertEqual(result[0]["Board"], 2)

    def test_partially_busy_boards(self):
        # Board 1 has only 1 ongoing game (Code allows < 2)
        existing = pd.DataFrame({
            'Board': [1],
            'Results White': [None],
            'Results Black': [None]
        })
        
        new_pairings = [{"White Player": 5}]
        result = assign_board_numbers(new_pairings, existing)
        
        # Can still assign to 1 because 1 < 2
        self.assertEqual(result[0]["Board"], 1)
    
    def test_wraps_around_max_boards(self):
        # We need to simulate usage filling up
        # Let's mock existing having no games, but ask for 35 pairings
        # (Assuming MAX_BOARDS is 30)
        
        new_pairings = [{"id": i} for i in range(MAX_BOARDS + 5)]
        result = assign_board_numbers(new_pairings, self.empty_existing)
        
        self.assertEqual(result[MAX_BOARDS - 1]["Board"], 30)
        self.assertEqual(result[MAX_BOARDS]["Board"], 1) # Wraps around


class TestCanPair(unittest.TestCase):
    """Test pairing eligibility logic."""
    
    def test_rematch_not_allowed(self):
        p1_state = {"opponents": {2}, "last_color": "White"}
        p2_state = {"opponents": {1}, "last_color": "Black"}
        self.assertFalse(can_pair(p1_state, p2_state, 2))
    
    def test_same_last_color_not_allowed(self):
        # Strict logic in source code: return color1 != color2
        p1_state = {"opponents": set(), "last_color": "White"}
        p2_state = {"opponents": set(), "last_color": "White"}
        self.assertFalse(can_pair(p1_state, p2_state, 2))
    
    def test_different_last_colors_allowed(self):
        p1_state = {"opponents": set(), "last_color": "White"}
        p2_state = {"opponents": set(), "last_color": "Black"}
        self.assertTrue(can_pair(p1_state, p2_state, 2))
    
    def test_one_new_player_allowed(self):
        p1_state = {"opponents": set(), "last_color": None}
        p2_state = {"opponents": set(), "last_color": "White"}
        self.assertTrue(can_pair(p1_state, p2_state, 2))


class TestAssignColors(unittest.TestCase):
    """Test color assignment logic."""
    
    @patch('random.choice')
    def test_both_new_players_random(self, mock_random):
        player_state = {
            1: {"last_color": None},
            2: {"last_color": None}
        }
        
        # Test Case A: random returns True -> (1, 2)
        mock_random.return_value = True
        self.assertEqual(assign_colors(1, 2, player_state), (1, 2))
        
        # Test Case B: random returns False -> (2, 1)
        mock_random.return_value = False
        self.assertEqual(assign_colors(1, 2, player_state), (2, 1))
    
    def test_new_vs_had_white(self):
        player_state = {
            1: {"last_color": None},
            2: {"last_color": "White"}
        }
        # Player 2 had White, so Player 2 must be Black -> (1, 2)
        # Logic in code: if color2 == "White" -> (p1, p2)
        self.assertEqual(assign_colors(1, 2, player_state), (1, 2))
    
    def test_white_then_black(self):
        player_state = {
            1: {"last_color": "White"},
            2: {"last_color": "Black"}
        }
        # P1 had White, needs Black. P2 had Black, needs White.
        # Result: P2 is White, P1 is Black -> (2, 1)
        self.assertEqual(assign_colors(1, 2, player_state), (2, 1))


class TestGeneratePairings(unittest.TestCase):
    """Test pairing generation logic."""
    
    def setUp(self):
        # Base template for a player
        self.base_player = {
            "currently_playing": False, "games_played": 0, 
            "points": 0.0, "opponents": set(), 
            "last_color": None, "name": "Player"
        }

    def test_no_players_waiting(self):
        player_state = {
            1: {**self.base_player, "currently_playing": True},
            2: {**self.base_player, "currently_playing": True}
        }
        self.assertEqual(generate_pairings(player_state), [])
    
    def test_odd_number_of_players(self):
        # 3 players available. Should produce 1 pair, leave 1 out.
        player_state = {
            1: self.base_player.copy(),
            2: self.base_player.copy(),
            3: self.base_player.copy()
        }
        pairings = generate_pairings(player_state)
        self.assertEqual(len(pairings), 1)
    
    def test_color_constraint_prevents_pairing(self):
        # P1 and P2 both had White. Strict rule says no pairing.
        player_state = {
            1: {**self.base_player, "last_color": "White", "games_played": 1},
            2: {**self.base_player, "last_color": "White", "games_played": 1}
        }
        pairings = generate_pairings(player_state)
        self.assertEqual(len(pairings), 0)

    def test_sorting_by_games_and_points(self):
        # P1: 2 games, 2 pts
        # P2: 2 games, 1 pts
        # P3: 1 game, 1 pt
        # P4: 1 game, 0 pts
        # Sorting logic in code: (games_played, -points)
        # Sort Order Expected: P3, P4, P1, P2 (Wait, actually P3(1), P4(1), P1(2), P2(2))
        
        player_state = {
            1: {**self.base_player, "games_played": 2, "points": 2.0, "last_color": "White"},
            2: {**self.base_player, "games_played": 2, "points": 1.0, "last_color": "Black"},
            3: {**self.base_player, "games_played": 1, "points": 1.0, "last_color": "White"},
            4: {**self.base_player, "games_played": 1, "points": 0.0, "last_color": "Black"},
        }
        
        pairings = generate_pairings(player_state)
        
        # Should pair based on games count first
        # P3 (White) vs P4 (Black) -> Valid
        # P1 (White) vs P2 (Black) -> Valid
        self.assertEqual(len(pairings), 2)
        
        # Verify P3 plays P4
        pair_sets = [{p["White Player"], p["Black Player"]} for p in pairings]
        self.assertIn({3, 4}, pair_sets)
        self.assertIn({1, 2}, pair_sets)


class TestCalculateStandings(unittest.TestCase):
    """Test standings calculation."""
    
    def test_buchholz_tiebreak(self):
        # Bob and Alice both have 1 pt.
        # Bob's opponents scored 1.0 (Higher Buchholz)
        # Alice's opponents scored 0.0 (Lower Buchholz)
        player_state = {
            1: {"name": "Alice", "points": 1.0, "opponent_scores": [0.0]},
            2: {"name": "Bob", "points": 1.0, "opponent_scores": [1.0]},
        }
        
        standings = calculate_standings(player_state)
        
        self.assertEqual(standings[0]["Player Name"], "Bob")
        self.assertEqual(standings[1]["Player Name"], "Alice")
        self.assertEqual(standings[0]["Pos"], 1)
        self.assertEqual(standings[1]["Pos"], 2)


class TestExcelOperations(unittest.TestCase):
    """Test Excel reading and writing operations using Mocks."""
    
    @patch('openpyxl.load_workbook')
    def test_append_pairings(self, mock_wb):
        from main import append_pairings_to_excel
        
        mock_sheet = MagicMock()
        mock_wb.return_value.__getitem__.return_value = mock_sheet
        
        new_pairings = [{"Round": 1, "Board": 1, "White Player": 1, "Black Player": 2}]
        player_state = {
            1: {"name": "Alice"},
            2: {"name": "Bob"}
        }
        
        append_pairings_to_excel("dummy.xlsx", new_pairings, player_state)
        
        # Verify append was called
        mock_sheet.append.assert_called_once()
        args = mock_sheet.append.call_args[0][0]
        # Check structure: Round, Board, W_ID, W_Name, B_ID, B_Name...
        self.assertEqual(args[0], 1) # Round
        self.assertEqual(args[3], "Alice") # W Name
        
        mock_wb.return_value.save.assert_called_with("dummy.xlsx")

    @patch('openpyxl.load_workbook')
    def test_write_standings(self, mock_wb):
        from main import write_standings_to_excel
        
        mock_wb_instance = mock_wb.return_value
        # Mock sheetnames to trigger deletion logic
        mock_wb_instance.sheetnames = ["Standings"]
        mock_sheet = MagicMock()
        mock_wb_instance.create_sheet.return_value = mock_sheet
        
        standings = [{"Pos": 1, "Player Name": "Alice", "Pt": 1.0, "BucT": 0, "Ber": 0}]
        
        write_standings_to_excel("dummy.xlsx", standings)
        
        # Verify old sheet deleted
        mock_wb_instance.__delitem__.assert_called_with("Standings")
        # Verify new sheet created
        mock_wb_instance.create_sheet.assert_called_with("Standings")
        # Verify header + 1 row appended
        self.assertEqual(mock_sheet.append.call_count, 2) 


if __name__ == "__main__":
    unittest.main()