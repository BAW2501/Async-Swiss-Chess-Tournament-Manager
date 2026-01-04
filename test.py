import unittest
import pandas as pd
from unittest.mock import patch, MagicMock
import sys
import os

# Import the main module (adjust path as needed)
# Assuming the main code is in swiss_tournament.py
sys.path.insert(0, os.path.dirname(__file__))
import main as st


class TestSwissTournament(unittest.TestCase):
    
    def setUp(self):
        """Set up test data that mimics the Excel structure."""
        self.attendees_data = pd.DataFrame({
            "ID": [0, 1, 2, 3, 4, 5, 6, 7],
            "First Name": ["Armin", "Máté", "Martin", "Bence", "Levente", "Dióssi", "Bence", "Sezgin"],
            "Last Name": ["Kozek", "Morassi", "Dudás", "Juhász", "Puha", "Csaba", "Matuz", "Rustamov"]
        })
        
        # Round 1 complete, Round 2 has some results
        self.pairings_r1_complete = pd.DataFrame({
            "Round": [1, 1, 1, 1, 2, 2],
            "Board": [1, 2, 3, 4, 1, 2],
            "White Player": [0, 2, 5, 6, 3, 7],
            "Black Player": [1, 3, 4, 7, 0, 5],
            "Results White": [1.0, 0.0, 1.0, 0.0, 1.0, 0.0],
            "Results Black": [0.0, 1.0, 0.0, 1.0, 0.0, 1.0]
        })
        
        # Round 1 complete, Round 2 partially complete
        self.pairings_r2_incomplete = pd.DataFrame({
            "Round": [1, 1, 1, 1, 2, 2, 2, 2],
            "Board": [1, 2, 3, 4, 1, 2, 3, 4],
            "White Player": [0, 2, 5, 6, 3, 7, 1, 4],
            "Black Player": [1, 3, 4, 7, 0, 5, 2, 6],
            "Results White": [1.0, 0.0, 1.0, 0.0, 1.0, 0.0, None, None],
            "Results Black": [0.0, 1.0, 0.0, 1.0, 0.0, 1.0, None, None]
        })
        
        # Empty pairings
        self.pairings_empty = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
    
    def test_build_player_state_empty(self):
        """Test player state building with no games played."""
        state = st.build_player_state(self.attendees_data, self.pairings_empty)
        
        self.assertEqual(len(state), 8)
        self.assertEqual(state[0]["name"], "Armin Kozek")
        self.assertEqual(state[0]["points"], 0.0)
        self.assertEqual(state[0]["games_played"], 0)
        self.assertIsNone(state[0]["last_color"])
        self.assertEqual(len(state[0]["opponents"]), 0)
    
    def test_build_player_state_with_games(self):
        """Test player state building with completed games."""
        state = st.build_player_state(self.attendees_data, self.pairings_r1_complete)
        
        # Player 0: Won R1 as White, Lost R2 as Black
        self.assertEqual(state[0]["points"], 1.0)
        self.assertEqual(state[0]["games_played"], 2)
        self.assertEqual(state[0]["last_color"], "Black")
        self.assertEqual(state[0]["opponents"], {1, 3})
        self.assertEqual(state[0]["color_history"], ["White", "Black"])
        
        # Player 3: Lost R1 as Black, Won R2 as White
        self.assertEqual(state[3]["points"], 1.0)
        self.assertEqual(state[3]["games_played"], 2)
        self.assertEqual(state[3]["last_color"], "White")
        self.assertEqual(state[3]["opponents"], {2, 0})
        self.assertEqual(state[3]["color_history"], ["Black", "White"])
    
    def test_build_player_state_incomplete_games(self):
        """Test that incomplete games are ignored."""
        state = st.build_player_state(self.attendees_data, self.pairings_r2_incomplete)
        
        # Player 1: Won R1, but R2 game is incomplete
        self.assertEqual(state[1]["points"], 0.0)
        self.assertEqual(state[1]["games_played"], 1)
        self.assertEqual(state[1]["last_color"], "Black")
        
        # Player 2: Lost R1, but R2 game is incomplete
        self.assertEqual(state[2]["points"], 0.0)
        self.assertEqual(state[2]["games_played"], 1)
    
    def test_get_lowest_incomplete_round_empty(self):
        """Test finding Rmin with no pairings."""
        rmin = st.get_lowest_incomplete_round(self.pairings_empty)
        self.assertEqual(rmin, 1)
    
    def test_get_lowest_incomplete_round_all_complete(self):
        """Test finding Rmin when all rounds are complete."""
        rmin = st.get_lowest_incomplete_round(self.pairings_r1_complete)
        self.assertEqual(rmin, 3)  # Next round after max complete
    
    def test_get_lowest_incomplete_round_partial(self):
        """Test finding Rmin with incomplete games in R2."""
        rmin = st.get_lowest_incomplete_round(self.pairings_r2_incomplete)
        self.assertEqual(rmin, 2)
    
    def test_get_players_in_pending_pairings(self):
        """Test finding players in unfinished games."""
        pending = st.get_players_in_pending_pairings(self.pairings_r2_incomplete)
        
        # Players 1, 2, 4, 6 are in incomplete games
        self.assertEqual(pending, {1, 2, 4, 6})
    
    def test_get_players_in_pending_pairings_empty(self):
        """Test pending players with no incomplete games."""
        pending = st.get_players_in_pending_pairings(self.pairings_r1_complete)
        self.assertEqual(pending, set())
    
    def test_would_force_triple_color(self):
        """Test triple color detection."""
        state = {
            0: {"color_history": ["White", "White"]},
            1: {"color_history": ["Black"]},
            2: {"color_history": []},
            3: {"color_history": ["White", "Black"]}
        }
        
        # Would be 3rd White in a row
        self.assertTrue(st.would_force_triple_color(state, 0, "White"))
        # Would not be 3rd White in a row
        self.assertFalse(st.would_force_triple_color(state, 0, "Black"))
        
        # Not enough history
        self.assertFalse(st.would_force_triple_color(state, 1, "Black"))
        self.assertFalse(st.would_force_triple_color(state, 2, "White"))
        
        # Alternating colors
        self.assertFalse(st.would_force_triple_color(state, 3, "White"))
    
    def test_can_pair_no_rematch(self):
        """Test pairing rules: no rematches."""
        p1_state = {"opponents": {2, 3}}
        p2_state = {"opponents": {1, 4}}
        
        # Can pair (haven't played)
        self.assertTrue(st.can_pair(p1_state, p2_state, 5))
        
        # Cannot pair (already played)
        self.assertFalse(st.can_pair(p1_state, p2_state, 2))
    
    def test_assign_colors_both_new(self):
        """Test color assignment for two new players."""
        state = {
            0: {"last_color": None, "color_history": []},
            1: {"last_color": None, "color_history": []}
        }
        
        # Should return random assignment (either order)
        white, black = st.assign_colors(0, 1, state)
        self.assertIn(white, [0, 1])
        self.assertIn(black, [0, 1])
        self.assertNotEqual(white, black)
    
    def test_assign_colors_alternation(self):
        """Test color assignment with alternation preference."""
        state = {
            0: {"last_color": "White", "color_history": ["White"]},
            1: {"last_color": "Black", "color_history": ["Black"]}
        }
        
        # Should alternate: 0 gets Black, 1 gets White
        white, black = st.assign_colors(0, 1, state)
        self.assertEqual(white, 1)
        self.assertEqual(black, 0)
    
    def test_assign_colors_triple_violation(self):
        """Test that triple-color violation returns None."""
        state = {
            0: {"last_color": "White", "color_history": ["White", "White"]},
            1: {"last_color": "White", "color_history": ["White", "White"]}
        }
        
        # Both would get 3rd White in a row - should return None
        result = st.assign_colors(0, 1, state)
        self.assertIsNone(result)
    
    def test_generate_pairings_first_round(self):
        """Test generating first round pairings."""
        state = st.build_player_state(self.attendees_data, self.pairings_empty)
        pending = set()
        rmin = 1
        
        pairings = st.generate_pairings(state, pending, rmin)
        
        # Should create 4 pairings (8 players)
        self.assertEqual(len(pairings), 4)
        
        # All pairings should be for Round 1
        for pairing in pairings:
            self.assertEqual(pairing["Round"], 1)
            self.assertIsNotNone(pairing["White Player"])
            self.assertIsNotNone(pairing["Black Player"])
    
    def test_generate_pairings_respects_window(self):
        """Test that pairings respect the two-round window."""
        state = st.build_player_state(self.attendees_data, self.pairings_r1_complete)
        pending = set()
        rmin = 3
        
        # All players have played 2 games (R1 and R2 complete)
        # With rmin=3, rmax=4, all players should be available
        pairings = st.generate_pairings(state, pending, rmin)
        
        # Should create pairings for R3
        self.assertGreater(len(pairings), 0)
        for pairing in pairings:
            self.assertEqual(pairing["Round"], 3)
    
    def test_generate_pairings_skips_pending_players(self):
        """Test that players in pending games are not paired."""
        state = st.build_player_state(self.attendees_data, self.pairings_r2_incomplete)
        pending = st.get_players_in_pending_pairings(self.pairings_r2_incomplete)
        rmin = 2
        
        # Players 1, 2, 4, 6 are pending
        pairings = st.generate_pairings(state, pending, rmin)
        
        # Check that pending players are not in new pairings
        for pairing in pairings:
            self.assertNotIn(pairing["White Player"], pending)
            self.assertNotIn(pairing["Black Player"], pending)
    
    def test_assign_boards_to_pairings_empty(self):
        """Test board assignment with no busy boards."""
        new_pairings = [
            {"Round": 1, "Board": None, "White Player": 0, "Black Player": 1},
            {"Round": 1, "Board": None, "White Player": 2, "Black Player": 3}
        ]
        
        assigned = st.assign_boards_to_pairings(new_pairings, self.pairings_empty)
        
        # Should assign boards 1 and 2
        self.assertEqual(assigned[0]["Board"], 1)
        self.assertEqual(assigned[1]["Board"], 2)
    
    def test_assign_boards_respects_busy_boards(self):
        """Test that board assignment avoids busy boards."""
        new_pairings = [
            {"Round": 2, "Board": None, "White Player": 4, "Black Player": 6}
        ]
        
        # Boards 1 and 2 have incomplete games
        incomplete_pairings = pd.DataFrame({
            "Round": [2, 2],
            "Board": [1, 2],
            "White Player": [1, 3],
            "Black Player": [2, 5],
            "Results White": [None, None],
            "Results Black": [None, None]
        })
        
        assigned = st.assign_boards_to_pairings(new_pairings, incomplete_pairings)
        
        # Should assign board 3 (first free board)
        self.assertEqual(assigned[0]["Board"], 3)
    
    def test_calculate_standings(self):
        """Test standings calculation with Buchholz."""
        state = st.build_player_state(self.attendees_data, self.pairings_r1_complete)
        standings = st.calculate_standings(state)
        
        # Should have 8 players
        self.assertEqual(len(standings), 8)
        
        # Check structure
        for standing in standings:
            self.assertIn("Pos", standing)
            self.assertIn("Player Name", standing)
            self.assertIn("Pt", standing)
            self.assertIn("BucT", standing)
        
        # Positions should be 1 to 8
        positions = [s["Pos"] for s in standings]
        self.assertEqual(sorted(positions), list(range(1, 9)))
        
        # Points should be sorted descending
        points = [s["Pt"] for s in standings]
        self.assertEqual(points, sorted(points, reverse=True))
    
    def test_calculate_standings_tie_break(self):
        """Test that Buchholz tie-break works correctly."""
        # Create specific scenario where two players have same points
        pairings = pd.DataFrame({
            "Round": [1, 1],
            "Board": [1, 2],
            "White Player": [0, 2],
            "Black Player": [1, 3],
            "Results White": [1.0, 1.0],
            "Results Black": [0.0, 0.0]
        })
        
        state = st.build_player_state(self.attendees_data, pairings)
        standings = st.calculate_standings(state)
        
        # Players 0 and 2 both have 1 point
        # Player 0 beat player 1 (0 pts), Player 2 beat player 3 (0 pts)
        # Buchholz should be equal
        p0_standing = next(s for s in standings if s["Player Name"] == "Armin Kozek")
        p2_standing = next(s for s in standings if s["Player Name"] == "Martin Dudás")
        
        self.assertEqual(p0_standing["Pt"], 1.0)
        self.assertEqual(p2_standing["Pt"], 1.0)
        self.assertEqual(p0_standing["BucT"], p2_standing["BucT"])
    
    @patch('openpyxl.load_workbook')
    def test_append_pairings_to_excel(self, mock_load_wb):
        """Test appending pairings to Excel."""
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_wb.__getitem__.return_value = mock_ws
        mock_load_wb.return_value = mock_wb
        
        state = st.build_player_state(self.attendees_data, self.pairings_empty)
        new_pairings = [
            {"Round": 1, "Board": 1, "White Player": 0, "Black Player": 1}
        ]
        
        st.append_pairings_to_excel("test.xlsx", new_pairings, state)
        
        # Check that append was called
        mock_ws.append.assert_called_once()
        call_args = mock_ws.append.call_args[0][0]
        
        self.assertEqual(call_args[0], 1)  # Round
        self.assertEqual(call_args[1], 1)  # Board
        self.assertEqual(call_args[2], 0)  # White Player ID
        self.assertEqual(call_args[4], 1)  # Black Player ID
        
        # Check that save was called
        mock_wb.save.assert_called_once_with("test.xlsx")
    
    @patch('openpyxl.load_workbook')
    def test_write_standings_to_excel(self, mock_load_wb):
        """Test writing standings to Excel."""
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_wb.create_sheet.return_value = mock_ws
        mock_wb.sheetnames = []
        mock_load_wb.return_value = mock_wb
        
        standings = [
            {"Pos": 1, "Player Name": "Test Player", "Pt": 1.5, "BucT": 2.0}
        ]
        
        st.write_standings_to_excel("test.xlsx", standings)
        
        # Check that headers were written
        calls = mock_ws.append.call_args_list
        self.assertEqual(calls[0][0][0], ["Pos", "Player Name", "Pt", "BucT"])
        
        # Check that data was written
        self.assertEqual(calls[1][0][0], [1, "Test Player", 1.5, 2.0])
        
        # Check that save was called
        mock_wb.save.assert_called_once_with("test.xlsx")


class TestEdgeCases(unittest.TestCase):
    """Test edge cases and error conditions."""
    
    def setUp(self):
        self.attendees_data = pd.DataFrame({
            "ID": [0, 1, 2],
            "First Name": ["Alice", "Bob", "Charlie"],
            "Last Name": ["A", "B", "C"]
        })
    
    def test_odd_number_of_players(self):
        """Test pairing with odd number of available players."""
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        state = st.build_player_state(self.attendees_data, pairings)
        pending = set()
        
        pairings = st.generate_pairings(state, pending, 1)
        
        # Should create 1 pairing, 1 player unpaired
        self.assertEqual(len(pairings), 1)
    
    def test_all_players_paired(self):
        """Test when all players are in pending games."""
        pairings = pd.DataFrame({
            "Round": [1],
            "Board": [1],
            "White Player": [0],
            "Black Player": [1],
            "Results White": [None],
            "Results Black": [None]
        })
        
        state = st.build_player_state(self.attendees_data, pairings)
        pending = st.get_players_in_pending_pairings(pairings)
        
        new_pairings = st.generate_pairings(state, pending, 1)
        
        # Player 2 is available, but can't pair alone
        self.assertEqual(len(new_pairings), 0)
    
    def test_all_boards_busy(self):
        """Test board assignment when all boards are busy."""
        # Create 30 incomplete pairings
        pairings_data = {
            "Round": [1] * 30,
            "Board": list(range(1, 31)),
            "White Player": list(range(0, 30)),
            "Black Player": list(range(30, 60)),
            "Results White": [None] * 30,
            "Results Black": [None] * 30
        }
        busy_pairings = pd.DataFrame(pairings_data)
        
        new_pairings = [
            {"Round": 1, "Board": None, "White Player": 60, "Black Player": 61}
        ]
        
        assigned = st.assign_boards_to_pairings(new_pairings, busy_pairings)
        
        # Should assign "?" when no boards available
        self.assertEqual(assigned[0]["Board"], "?")


if __name__ == "__main__":
    # Run tests with verbose output
    unittest.main(verbosity=2)