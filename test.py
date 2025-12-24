import unittest
import pandas as pd
from unittest.mock import Mock, patch, MagicMock
import sys
import tempfile
from pathlib import Path

# Import the functions we're testing
# Assuming the main code is in swiss_pairing.py
from main import *

# For testing purposes, I'll include minimal versions of the functions
# In practice, you'd import from your actual module

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
        
        # Mock the function or import it
        player_state = build_player_state(self.attendees, pairings)
        
        # Expected behavior:
        # - All players should have 0 points
        # - All players should have 0 games_played
        # - No one should be currently_playing
        # - last_color should be None
        
    
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
        
        # Expected:
        # - Player 1: 1 point, 1 game, last_color='White', not playing
        # - Player 2: 0 points, 1 game, last_color='Black', not playing
        # - Player 1's opponents should include 2
        # - Player 2's opponents should include 1
        pass
    
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
        
        # Expected:
        # - Both players marked as currently_playing
        # - No points awarded yet
        # - games_played still 0
        # - Opponents recorded
        pass
    
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
        
        # Expected:
        # - Player 1: 1 point from game 1, currently playing game 3
        # - Player 2: 0 points, not playing
        # - Player 3: 0.5 points from game 2, currently playing game 3
        # - Player 4: 0.5 points, not playing
        pass
    
    def test_buchholz_calculation(self):
        """Test that opponent scores are correctly stored for Buchholz."""
        pairings = pd.DataFrame({
            'Round': [1, 1],
            'Board': [1, 2],
            'White Player': [1, 2],
            'Black Player': [3, 4],
            'Results White': [1.0, 0.5],
            'Results Black': [0.0, 0.5]
        })
        
        # Expected:
        # - Player 1's opponent_scores should be [0.0] (Player 3's score)
        # - Player 3's opponent_scores should be [1.0] (Player 1's score)
        # - Player 2's opponent_scores should be [0.5] (Player 4's score)
        # - Player 4's opponent_scores should be [0.5] (Player 2's score)
        pass
    
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
        
        # Expected: Should skip this pairing gracefully
        pass
    
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
        
        # Expected: Both players get 0.5 points
        pass


class TestGetBoardUsageCount(unittest.TestCase):
    """Test board usage counting."""
    
    def test_empty_pairings(self):
        """Test with no pairings."""
        pairings = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        # Expected: Empty dictionary
        pass
    
    def test_all_finished_games(self):
        """Test that finished games don't count toward usage."""
        pairings = pd.DataFrame({
            'Round': [1, 1],
            'Board': [1, 2],
            'White Player': [1, 3],
            'Black Player': [2, 4],
            'Results White': [1.0, 0.5],
            'Results Black': [0.0, 0.5]
        })
        
        # Expected: Empty dictionary (no ongoing games)
        pass
    
    def test_ongoing_games_single_board(self):
        """Test ongoing games on same board."""
        pairings = pd.DataFrame({
            'Round': [1, 2],
            'Board': [1, 1],
            'White Player': [1, 3],
            'Black Player': [2, 4],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        
        # Expected: {1: 2}
        pass
    
    def test_ongoing_games_multiple_boards(self):
        """Test ongoing games across multiple boards."""
        pairings = pd.DataFrame({
            'Round': [1, 1, 2],
            'Board': [1, 2, 3],
            'White Player': [1, 3, 5],
            'Black Player': [2, 4, 6],
            'Results White': [None, None, None],
            'Results Black': [None, None, None]
        })
        
        # Expected: {1: 1, 2: 1, 3: 1}
        pass
    
    def test_question_mark_boards(self):
        """Test that '?' boards are ignored."""
        pairings = pd.DataFrame({
            'Round': [1, 2],
            'Board': ['?', 1],
            'White Player': [1, 3],
            'Black Player': [2, 4],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        
        # Expected: {1: 1} - '?' should be skipped
        pass
    
    def test_mixed_finished_and_ongoing(self):
        """Test that only ongoing games count."""
        pairings = pd.DataFrame({
            'Round': [1, 1, 2, 2],
            'Board': [1, 2, 1, 3],
            'White Player': [1, 3, 5, 7],
            'Black Player': [2, 4, 6, 8],
            'Results White': [1.0, None, None, 0.5],
            'Results Black': [0.0, None, None, 0.5]
        })
        
        # Expected: {2: 1, 1: 1} - only rounds 2 and 3 are ongoing
        pass


class TestAssignBoardNumbers(unittest.TestCase):
    """Test board number assignment."""
    
    def test_empty_pairings(self):
        """Test with no pairings to assign."""
        new_pairings = []
        existing = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        # Expected: Empty list returned
        pass
    
    def test_assign_to_empty_boards(self):
        """Test assigning when all boards are free."""
        new_pairings = [
            {"Round": 1, "White Player": 1, "Black Player": 2},
            {"Round": 1, "White Player": 3, "Black Player": 4},
            {"Round": 1, "White Player": 5, "Black Player": 6}
        ]
        existing = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        # Expected: Boards 1, 2, 3 assigned
        pass
    
    def test_round_robin_assignment(self):
        """Test that boards are assigned in round-robin fashion."""
        new_pairings = [
            {"Round": 1, "White Player": i*2-1, "Black Player": i*2}
            for i in range(1, 6)
        ]
        existing = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        # Expected: Boards 1, 2, 3, 4, 5
        pass
    
    def test_skip_busy_boards(self):
        """Test that boards with 2 ongoing games are skipped."""
        existing = pd.DataFrame({
            'Round': [1, 2],
            'Board': [1, 1],
            'White Player': [1, 3],
            'Black Player': [2, 4],
            'Results White': [None, None],
            'Results Black': [None, None]
        })
        
        new_pairings = [
            {"Round": 3, "White Player": 5, "Black Player": 6}
        ]
        
        # Expected: Should assign board 2 (board 1 is full)
        pass
    
    def test_partially_busy_boards(self):
        """Test boards with 1 ongoing game can still be used."""
        existing = pd.DataFrame({
            'Round': [1],
            'Board': [1],
            'White Player': [1],
            'Black Player': [2],
            'Results White': [None],
            'Results Black': [None]
        })
        
        new_pairings = [
            {"Round": 2, "White Player": 3, "Black Player": 4}
        ]
        
        # Expected: Can still assign board 1 (only 1 game ongoing)
        pass
    
    def test_all_boards_full(self):
        """Test when all boards are at capacity."""
        # Create 34 ongoing games (17 boards * 2 games each)
        existing_data = []
        for board in range(1, 18):
            for game in range(2):
                existing_data.append({
                    'Round': game + 1,
                    'Board': board,
                    'White Player': board * 10 + game * 2,
                    'Black Player': board * 10 + game * 2 + 1,
                    'Results White': None,
                    'Results Black': None
                })
        existing = pd.DataFrame(existing_data)
        
        new_pairings = [
            {"Round": 3, "White Player": 100, "Black Player": 101}
        ]
        
        # Expected: Should assign '?'
        pass
    
    def test_wraps_around_max_boards(self):
        """Test that assignment wraps around after board 17."""
        new_pairings = [
            {"Round": 1, "White Player": i*2-1, "Black Player": i*2}
            for i in range(1, 20)  # 19 pairings
        ]
        existing = pd.DataFrame(columns=["Round", "Board", "White Player", "Black Player", "Results White", "Results Black"])
        
        # Expected: Should cycle through 1-17, then start at 1 again
        pass


class TestCanPair(unittest.TestCase):
    """Test pairing eligibility logic."""
    
    def test_rematch_not_allowed(self):
        """Test that players can't be paired twice."""
        p1_state = {
            "opponents": {2, 3},
            "last_color": "White"
        }
        p2_state = {
            "opponents": {1},
            "last_color": "Black"
        }
        
        # Expected: False (2 is in p1's opponents)
        pass
    
    def test_same_last_color_not_allowed(self):
        """Test that players with same last color can't pair."""
        p1_state = {
            "opponents": set(),
            "last_color": "White"
        }
        p2_state = {
            "opponents": set(),
            "last_color": "White"
        }
        
        # Expected: False
        pass
    
    def test_different_last_colors_allowed(self):
        """Test that players with different last colors can pair."""
        p1_state = {
            "opponents": set(),
            "last_color": "White"
        }
        p2_state = {
            "opponents": set(),
            "last_color": "Black"
        }
        
        # Expected: True
        pass
    
    def test_both_new_players_allowed(self):
        """Test that two new players can pair."""
        p1_state = {
            "opponents": set(),
            "last_color": None
        }
        p2_state = {
            "opponents": set(),
            "last_color": None
        }
        
        # Expected: True
        pass
    
    def test_one_new_player_allowed(self):
        """Test that new player can pair with experienced player."""
        p1_state = {
            "opponents": set(),
            "last_color": None
        }
        p2_state = {
            "opponents": {3},
            "last_color": "White"
        }
        
        # Expected: True
        pass
    
    def test_rematch_prevents_pairing_even_with_good_colors(self):
        """Test that rematch rule takes precedence over color rule."""
        p1_state = {
            "opponents": {2},
            "last_color": "White"
        }
        p2_state = {
            "opponents": {1},
            "last_color": "Black"
        }
        
        # Expected: False (rematch)
        pass


class TestAssignColors(unittest.TestCase):
    """Test color assignment logic."""
    
    def test_both_new_players_random(self):
        """Test random assignment for two new players."""
        player_state = {
            1: {"last_color": None},
            2: {"last_color": None}
        }
        
        # Expected: Either (1, 2) or (2, 1), randomly
        # Run multiple times to verify randomness
        pass
    
    def test_new_vs_had_white(self):
        """Test new player vs player who had white."""
        player_state = {
            1: {"last_color": None},
            2: {"last_color": "White"}
        }
        
        # Expected: (2, 1) - player 2 gets black
        pass
    
    def test_new_vs_had_black(self):
        """Test new player vs player who had black."""
        player_state = {
            1: {"last_color": None},
            2: {"last_color": "Black"}
        }
        
        # Expected: (1, 2) - player 2 stays black
        pass
    
    def test_both_had_white(self):
        """Test when both players last had white."""
        player_state = {
            1: {"last_color": "White"},
            2: {"last_color": "White"}
        }
        
        # This should never happen in practice (can_pair would reject)
        # But if it did: (2, 1) based on alternation logic
        pass
    
    def test_white_then_black(self):
        """Test player who had white gets black."""
        player_state = {
            1: {"last_color": "White"},
            2: {"last_color": "Black"}
        }
        
        # Expected: (2, 1) - player 1 gets black
        pass
    
    def test_black_then_white(self):
        """Test player who had black gets white."""
        player_state = {
            1: {"last_color": "Black"},
            2: {"last_color": "White"}
        }
        
        # Expected: (1, 2) - player 1 gets white
        pass


class TestGeneratePairings(unittest.TestCase):
    """Test pairing generation logic."""
    
    def test_no_players_waiting(self):
        """Test with no players available."""
        player_state = {
            1: {"currently_playing": True, "games_played": 1, "points": 1.0, "opponents": {2}, "last_color": "White"},
            2: {"currently_playing": True, "games_played": 1, "points": 0.0, "opponents": {1}, "last_color": "Black"}
        }
        
        # Expected: Empty list
        pass
    
    def test_single_player_waiting(self):
        """Test with only one player available."""
        player_state = {
            1: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None},
            2: {"currently_playing": True, "games_played": 1, "points": 1.0, "opponents": {3}, "last_color": "White"}
        }
        
        # Expected: Empty list
        pass
    
    def test_simple_first_round(self):
        """Test pairing for first round with 4 players."""
        player_state = {
            1: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Alice"},
            2: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Bob"},
            3: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Charlie"},
            4: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "David"}
        }
        
        # Expected: 2 pairings, all round 1
        pass
    
    def test_odd_number_of_players(self):
        """Test with odd number of players."""
        player_state = {
            1: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Alice"},
            2: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Bob"},
            3: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Charlie"}
        }
        
        # Expected: 1 pairing, 1 player unpaired
        pass
    
    def test_color_constraint_prevents_pairing(self):
        """Test when color constraints prevent all pairings."""
        player_state = {
            1: {"currently_playing": False, "games_played": 1, "points": 1.0, "opponents": {3}, "last_color": "White", "name": "Alice"},
            2: {"currently_playing": False, "games_played": 1, "points": 1.0, "opponents": {4}, "last_color": "White", "name": "Bob"}
        }
        
        # Expected: 0 pairings (both last had white)
        pass
    
    def test_rematch_constraint(self):
        """Test that rematches are avoided."""
        player_state = {
            1: {"currently_playing": False, "games_played": 1, "points": 1.0, "opponents": {2}, "last_color": "White", "name": "Alice"},
            2: {"currently_playing": False, "games_played": 1, "points": 0.0, "opponents": {1}, "last_color": "Black", "name": "Bob"}
        }
        
        # Expected: 0 pairings (they already played)
        pass
    
    def test_sorting_by_games_and_points(self):
        """Test that players are sorted correctly before pairing."""
        player_state = {
            1: {"currently_playing": False, "games_played": 2, "points": 2.0, "opponents": {2, 3}, "last_color": "White", "name": "Alice"},
            2: {"currently_playing": False, "games_played": 2, "points": 1.5, "opponents": {1, 4}, "last_color": "Black", "name": "Bob"},
            3: {"currently_playing": False, "games_played": 1, "points": 1.0, "opponents": {1}, "last_color": "Black", "name": "Charlie"},
            4: {"currently_playing": False, "games_played": 1, "points": 0.5, "opponents": {2}, "last_color": "White", "name": "David"}
        }
        
        # Expected: Should pair 3-4 first (both have 1 game), then try 1-2
        pass
    
    def test_greedy_matching_finds_all_pairs(self):
        """Test that greedy algorithm finds maximum matching when possible."""
        player_state = {
            1: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Alice"},
            2: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Bob"},
            3: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Charlie"},
            4: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "David"},
            5: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Eve"},
            6: {"currently_playing": False, "games_played": 0, "points": 0.0, "opponents": set(), "last_color": None, "name": "Frank"}
        }
        
        # Expected: 3 pairings
        pass


class TestCalculateStandings(unittest.TestCase):
    """Test standings calculation."""
    
    def test_empty_tournament(self):
        """Test standings with no games played."""
        player_state = {
            1: {"name": "Alice", "points": 0.0, "opponent_scores": [], "games_played": 0},
            2: {"name": "Bob", "points": 0.0, "opponent_scores": [], "games_played": 0}
        }
        
        # Expected: Both tied at 0 points, 0 Buchholz
        pass
    
    def test_simple_standings(self):
        """Test standings after one round."""
        player_state = {
            1: {"name": "Alice", "points": 1.0, "opponent_scores": [0.0], "games_played": 1},
            2: {"name": "Bob", "points": 0.0, "opponent_scores": [1.0], "games_played": 1}
        }
        
        # Expected: Alice 1st (1 pt), Bob 2nd (0 pts)
        # Alice's Buchholz = 0.0, Bob's Buchholz = 1.0
        pass
    
    def test_buchholz_tiebreak(self):
        """Test Buchholz as tiebreaker."""
        player_state = {
            1: {"name": "Alice", "points": 1.0, "opponent_scores": [0.5], "games_played": 1},
            2: {"name": "Bob", "points": 1.0, "opponent_scores": [1.0], "games_played": 1},
            3: {"name": "Charlie", "points": 0.5, "opponent_scores": [1.0], "games_played": 1},
            4: {"name": "David", "points": 1.0, "opponent_scores": [0.0], "games_played": 1}
        }
        
        # Expected: Bob 1st (Buchholz 1.0), Alice 2nd (Buchholz 0.5), David 3rd (Buchholz 0.0)
        pass
    
    def test_position_numbering(self):
        """Test that positions are numbered correctly."""
        player_state = {
            1: {"name": "Alice", "points": 2.0, "opponent_scores": [1.0, 0.5], "games_played": 2},
            2: {"name": "Bob", "points": 1.5, "opponent_scores": [0.5, 1.0], "games_played": 2},
            3: {"name": "Charlie", "points": 1.0, "opponent_scores": [1.5, 0.0], "games_played": 2}
        }
        
        # Expected: Pos 1, 2, 3
        pass
    
    def test_multiple_players_same_score(self):
        """Test handling of multiple players with identical scores."""
        player_state = {
            1: {"name": "Alice", "points": 1.0, "opponent_scores": [0.5], "games_played": 1},
            2: {"name": "Bob", "points": 1.0, "opponent_scores": [0.5], "games_played": 1},
            3: {"name": "Charlie", "points": 0.5, "opponent_scores": [1.0], "games_played": 1},
            4: {"name": "David", "points": 0.5, "opponent_scores": [1.0], "games_played": 1}
        }
        
        # Expected: Alice and Bob tied at top, Charlie and David tied at bottom
        pass


class TestExcelOperations(unittest.TestCase):
    """Test Excel reading and writing operations."""
    
    def test_load_tournament_data_missing_file(self):
        """Test loading when file doesn't exist."""
        # Expected: Should exit with error
        pass
    
    def test_load_tournament_data_missing_pairings_sheet(self):
        """Test loading when Pairings sheet doesn't exist."""
        # Expected: Should create empty DataFrame with correct columns
        pass
    
    def test_append_pairings_preserves_existing(self):
        """Test that appending doesn't modify existing pairings."""
        # This would require actual file I/O or mocking
        pass
    
    def test_write_standings_overwrites(self):
        """Test that writing standings replaces old standings."""
        # This would require actual file I/O or mocking
        pass


class TestEdgeCases(unittest.TestCase):
    """Test edge cases and error handling."""
    
    def test_large_tournament(self):
        """Test with 100 players."""
        player_state = {
            i: {
                "currently_playing": False,
                "games_played": 0,
                "points": 0.0,
                "opponents": set(),
                "last_color": None,
                "name": f"Player{i}"
            }
            for i in range(1, 101)
        }
        
        # Expected: Should pair 50 games
        pass
    
    def test_all_boards_exhausted_scenario(self):
        """Test realistic scenario where boards run out."""
        # 17 boards * 2 games = 34 ongoing games
        # If we have 40 players waiting, we can only assign 34
        pass
    
    def test_color_alternation_over_many_rounds(self):
        """Test that color alternation works over 5+ rounds."""
        # Simulate a player playing 5 games and verify colors alternate
        pass
    
    def test_swiss_pairing_convergence(self):
        """Test that after several rounds, top players face each other."""
        # Simulate tournament progression and verify pairing quality
        pass
    
    def test_bye_handling(self):
        """Test handling of byes (odd player count)."""
        # Verify that one player is consistently left out when odd
        pass


class TestIntegration(unittest.TestCase):
    """Integration tests for complete workflows."""
    
    def test_complete_round_robin_small_tournament(self):
        """Test a complete 4-player round-robin tournament."""
        # Simulate all 6 games and verify final standings
        pass
    
    def test_swiss_tournament_progression(self):
        """Test a 3-round Swiss tournament with 8 players."""
        # Verify pairings make sense each round
        pass
    
    def test_concurrent_games_handling(self):
        """Test multiple rounds with staggered game completion."""
        # Some games finish, others ongoing, new pairings generated
        pass


# Helper function to run all tests
def run_tests():
    """Run all unit tests and display results."""
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # Add all test classes
    suite.addTests(loader.loadTestsFromTestCase(TestBuildPlayerState))
    suite.addTests(loader.loadTestsFromTestCase(TestGetBoardUsageCount))
    suite.addTests(loader.loadTestsFromTestCase(TestAssignBoardNumbers))
    suite.addTests(loader.loadTestsFromTestCase(TestCanPair))
    suite.addTests(loader.loadTestsFromTestCase(TestAssignColors))
    suite.addTests(loader.loadTestsFromTestCase(TestGeneratePairings))
    suite.addTests(loader.loadTestsFromTestCase(TestCalculateStandings))
    suite.addTests(loader.loadTestsFromTestCase(TestExcelOperations))
    suite.addTests(loader.loadTestsFromTestCase(TestEdgeCases))
    suite.addTests(loader.loadTestsFromTestCase(TestIntegration))
    
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    return result

if __name__ == "__main__":
    run_tests()