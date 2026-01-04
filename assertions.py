from main import load_tournament_data,build_player_state


if __name__ == "__main__":
    attendees, pairings = load_tournament_data()

    print(f"Loaded {len(attendees)} attendees and {len(pairings)} pairings.")

    # Build player states
    player_state = build_player_state(attendees, pairings)
    sorted_player_state = sorted(player_state.items(),key=lambda x: x[1]['points'],reverse=True)
    for player_id, state in sorted_player_state:
        print(f"Player {player_id}: {state}")