"""
Swiss Tournament Live Display Server
Watches tournament.xlsx and displays live standings and pairings
Built by BAW2501 - https://github.com/BAW2501
"""

from flask import Flask, render_template_string
import pandas as pd
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import threading
import time
from datetime import datetime

app = Flask(__name__)

# Global data storage
tournament_data = {
    'standings': [],
    'pairings': [],
    'last_update': None
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🏆 Amateur Chess Tournament - Live</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #D4A574 0%, #8B6F47 50%, #D4A574 100%);
            min-height: 100vh;
            padding: 15px;
            color: #2c1810;
        }

        .container {
            max-width: 1600px;
            margin: 0 auto;
        }

        .header {
            text-align: center;
            margin-bottom: 25px;
            background: rgba(255, 255, 255, 0.95);
            padding: 20px;
            border-radius: 20px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
            border: 3px solid #5d4e37;
        }

        .logo-container {
            margin: 15px 0;
        }

        .logo-container img {
            max-width: 500px;
            width: 100%;
            height: auto;
            border-radius: 15px;
            box-shadow: 0 5px 20px rgba(0, 0, 0, 0.2);
        }

        h1 {
            font-size: 2.5em;
            color: #5d4e37;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1);
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
        }

        .subtitle {
            font-size: 1.2em;
            color: #8B6F47;
            font-weight: 600;
            margin-top: 10px;
        }

        .last-update {
            margin-top: 12px;
            font-size: 0.9em;
            color: #666;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }

        /* TABS */
        .tabs {
            display: flex;
            gap: 10px;
            justify-content: center;
            margin-bottom: 25px;
        }

        .tab-button {
            background: rgba(255, 255, 255, 0.9);
            border: 3px solid #5d4e37;
            padding: 15px 40px;
            font-size: 1.3em;
            font-weight: bold;
            color: #5d4e37;
            cursor: pointer;
            border-radius: 15px;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            gap: 10px;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        }

        .tab-button:hover {
            transform: translateY(-3px);
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
        }

        .tab-button.active {
            background: linear-gradient(135deg, #5d4e37, #8B6F47);
            color: white;
            border-color: #8B6F47;
        }

        .tab-content {
            display: none;
        }

        .tab-content.active {
            display: block;
        }

        .section {
            background: rgba(255, 255, 255, 0.95);
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.2);
            border: 2px solid #5d4e37;
        }

        .section-title {
            font-size: 1.8em;
            color: #5d4e37;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 12px;
            border-bottom: 3px solid #D4A574;
            padding-bottom: 10px;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
        }

        thead {
            background: linear-gradient(135deg, #5d4e37 0%, #8B6F47 100%);
            color: white;
        }

        th {
            padding: 12px 10px;
            text-align: left;
            font-weight: 700;
            font-size: 0.95em;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }

        th i {
            margin-right: 6px;
        }

        td {
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
            font-size: 0.95em;
        }

        tr:hover {
            background-color: #fff9f0;
            transition: all 0.3s ease;
        }

        tr:last-child td {
            border-bottom: none;
        }

        .rank {
            font-weight: bold;
            color: #5d4e37;
            font-size: 1.1em;
        }

        .rank-1 {
            color: #FFD700;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
        }

        .rank-2 {
            color: #C0C0C0;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
        }

        .rank-3 {
            color: #CD7F32;
            text-shadow: 1px 1px 2px rgba(0, 0, 0, 0.3);
        }

        .points {
            font-weight: bold;
            color: #2c5f2d;
            font-size: 1.1em;
        }

        .board-number {
            background: linear-gradient(135deg, #5d4e37, #8B6F47);
            color: white;
            padding: 6px 12px;
            border-radius: 8px;
            font-weight: bold;
            display: inline-block;
            min-width: 50px;
            text-align: center;
            font-size: 0.9em;
        }

        .player-cell {
            font-weight: 600;
            color: #2c1810;
        }

        .vs-separator {
            text-align: center;
            color: #8B6F47;
            font-weight: bold;
            font-size: 1.1em;
        }

        .round-badge {
            background: #D4A574;
            color: #2c1810;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: bold;
            display: inline-block;
        }

        .result {
            font-weight: bold;
            padding: 4px 8px;
            border-radius: 5px;
            display: inline-block;
            min-width: 45px;
            text-align: center;
            font-size: 0.9em;
        }

        .result-win {
            background: #d4edda;
            color: #155724;
        }

        .result-draw {
            background: #fff3cd;
            color: #856404;
        }

        .result-loss {
            background: #f8d7da;
            color: #721c24;
        }

        .result-pending {
            background: #e2e3e5;
            color: #383d41;
        }

        .footer {
            text-align: center;
            margin-top: 30px;
            padding: 15px;
            background: rgba(255, 255, 255, 0.9);
            border-radius: 15px;
            border: 2px solid #5d4e37;
        }

        .footer a {
            color: #5d4e37;
            text-decoration: none;
            font-weight: bold;
            font-size: 1.05em;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: all 0.3s ease;
        }

        .footer a:hover {
            color: #D4A574;
            transform: translateY(-2px);
        }

        .empty-state {
            text-align: center;
            padding: 40px;
            color: #666;
            font-size: 1.2em;
        }

        .empty-state i {
            font-size: 3em;
            color: #D4A574;
            margin-bottom: 15px;
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }

        .live-indicator {
            display: inline-block;
            width: 12px;
            height: 12px;
            background: #ff4444;
            border-radius: 50%;
            animation: pulse 2s infinite;
            margin-right: 8px;
        }

        @media (max-width: 768px) {
            h1 {
                font-size: 1.8em;
            }

            .section {
                padding: 15px;
            }

            table {
                font-size: 0.85em;
            }

            th, td {
                padding: 8px 6px;
            }

            .tab-button {
                padding: 12px 25px;
                font-size: 1.1em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>
                <i class="fas fa-chess"></i>
                Amateur Chess Tournament
                <i class="fas fa-trophy"></i>
            </h1>
            <div class="subtitle">
                <span class="live-indicator"></span>
                LIVE STANDINGS & PAIRINGS
            </div>
            
            <div class="logo-container">
                <img src="/static/tournament_poster.jpg" alt="Tournament Poster" onerror="this.style.display='none'">
            </div>
            
            {% if last_update %}
            <div class="last-update">
                <i class="far fa-clock"></i>
                Last Updated: {{ last_update }}
            </div>
            {% endif %}
        </div>

        <!-- TABS -->
        <div class="tabs">
            <button class="tab-button active" onclick="showTab('standings')">
                <i class="fas fa-ranking-star"></i>
                Standings
            </button>
            <button class="tab-button" onclick="showTab('pairings')">
                <i class="fas fa-chess-board"></i>
                Pairings
            </button>
        </div>

        <!-- STANDINGS TAB -->
        <div id="standings-tab" class="tab-content active">
            <div class="section">
                <h2 class="section-title">
                    <i class="fas fa-ranking-star"></i>
                    Current Standings
                </h2>
                {% if standings %}
                <table>
                    <thead>
                        <tr>
                            <th><i class="fas fa-hashtag"></i> Rank</th>
                            <th><i class="fas fa-user"></i> Player</th>
                            <th><i class="fas fa-star"></i> Points</th>
                            <th><i class="fas fa-chart-line"></i> Buchholz</th>
                            <th><i class="fas fa-handshake"></i> Direct</th>
                            <th><i class="fas fa-calculator"></i> Berger</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for player in standings %}
                        <tr>
                            <td class="rank {% if player.Pos == 1 %}rank-1{% elif player.Pos == 2 %}rank-2{% elif player.Pos == 3 %}rank-3{% endif %}">
                                {% if player.Pos == 1 %}
                                    <i class="fas fa-crown"></i>
                                {% elif player.Pos == 2 %}
                                    <i class="fas fa-medal"></i>
                                {% elif player.Pos == 3 %}
                                    <i class="fas fa-award"></i>
                                {% endif %}
                                {{ player.Pos }}
                            </td>
                            <td class="player-cell">{{ player['Player Name'] }}</td>
                            <td class="points">{{ player.Pt }}</td>
                            <td>{{ player.BucT }}</td>
                            <td>{{ player.DE }}</td>
                            <td>{{ player.Ber }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% else %}
                <div class="empty-state">
                    <i class="fas fa-hourglass-start"></i>
                    <p>Standings will appear here once games are completed</p>
                </div>
                {% endif %}
            </div>
        </div>

        <!-- PAIRINGS TAB -->
        <div id="pairings-tab" class="tab-content">
            <div class="section">
                <h2 class="section-title">
                    <i class="fas fa-chess-board"></i>
                    Current Pairings
                </h2>
                {% if pairings %}
                <table>
                    <thead>
                        <tr>
                            <th><i class="fas fa-layer-group"></i> Round</th>
                            <th><i class="fas fa-chess-board"></i> Board</th>
                            <th><i class="fas fa-chess-king"></i> White Player</th>
                            <th></th>
                            <th><i class="fas fa-chess-queen"></i> Black Player</th>
                            <th><i class="fas fa-trophy"></i> Result</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for pairing in pairings %}
                        <tr>
                            <td><span class="round-badge">R{{ pairing.Round }}</span></td>
                            <td><span class="board-number">{{ pairing.Board }}</span></td>
                            <td class="player-cell">
                                <i class="fas fa-square" style="color: white; text-shadow: 0 0 1px black;"></i>
                                {{ pairing['White Name'] }}
                            </td>
                            <td class="vs-separator">VS</td>
                            <td class="player-cell">
                                <i class="fas fa-square" style="color: black;"></i>
                                {{ pairing['Black Name'] }}
                            </td>
                            <td>
                                {% if pairing.result_status == 'White Win' %}
                                    <span class="result result-win">1-0</span>
                                {% elif pairing.result_status == 'Black Win' %}
                                    <span class="result result-loss">0-1</span>
                                {% elif pairing.result_status == 'Draw' %}
                                    <span class="result result-draw">½-½</span>
                                {% else %}
                                    <span class="result result-pending"><i class="fas fa-hourglass-half"></i> Live</span>
                                {% endif %}
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                {% else %}
                <div class="empty-state">
                    <i class="fas fa-chess-knight"></i>
                    <p>No pairings yet. Run the pairing script to generate matches!</p>
                </div>
                {% endif %}
            </div>
        </div>

        <div class="footer">
            <p>
                <i class="fas fa-code"></i>
                Built with ♟️ by 
                <a href="https://github.com/BAW2501" target="_blank">
                    <i class="fab fa-github"></i>
                    BAW2501
                </a>
            </p>
        </div>
    </div>

    <script>
        // Tab switching
        function showTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.classList.remove('active');
            });
            
            // Remove active from all buttons
            document.querySelectorAll('.tab-button').forEach(btn => {
                btn.classList.remove('active');
            });
            
            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            
            // Activate button
            event.target.closest('.tab-button').classList.add('active');
        }

        // Auto-refresh every 10 seconds
        setTimeout(function() {
            location.reload();
        }, 10000);
    </script>
</body>
</html>
"""

class ExcelFileHandler(FileSystemEventHandler):
    """Handles file system events for the Excel file."""
    
    def __init__(self, filepath):
        self.filepath = filepath
        self.last_modified = 0
        
    def on_modified(self, event):
        if event.src_path.endswith('tournament.xlsx'):
            # Debounce: only reload if at least 2 seconds have passed
            current_time = time.time()
            if current_time - self.last_modified > 2:
                self.last_modified = current_time
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Excel file changed - reloading data...")
                time.sleep(0.5)  # Give Excel time to finish writing
                load_tournament_data()

def load_tournament_data():
    """Load data from Excel file."""
    try:
        filepath = "tournament.xlsx"
        
        # Load standings
        try:
            standings_df = pd.read_excel(filepath, sheet_name="Standings", engine='openpyxl')
            tournament_data['standings'] = standings_df.to_dict('records')
        except Exception as e:
            print(f"Could not load standings: {e}")
            tournament_data['standings'] = []
        
        # Load pairings
        try:
            pairings_df = pd.read_excel(filepath, sheet_name="Pairings", engine='openpyxl')
            
            # Add result status
            pairings_list = []
            for _, row in pairings_df.iterrows():
                pairing = row.to_dict()
                
                # Determine result status
                if pd.notna(row.get('Results White')) and pd.notna(row.get('Results Black')):
                    white_result = row['Results White']
                    black_result = row['Results Black']
                    
                    if white_result == 1:
                        pairing['result_status'] = 'White Win'
                    elif black_result == 1:
                        pairing['result_status'] = 'Black Win'
                    elif white_result == 0.5:
                        pairing['result_status'] = 'Draw'
                    else:
                        pairing['result_status'] = 'Pending'
                else:
                    pairing['result_status'] = 'Pending'
                
                pairings_list.append(pairing)
            
            tournament_data['pairings'] = pairings_list
        except Exception as e:
            print(f"Could not load pairings: {e}")
            tournament_data['pairings'] = []
        
        tournament_data['last_update'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Data loaded: {len(tournament_data['standings'])} standings, {len(tournament_data['pairings'])} pairings")
        
    except Exception as e:
        print(f"Error loading tournament data: {e}")

@app.route('/')
def index():
    """Main page route."""
    return render_template_string(
        HTML_TEMPLATE,
        standings=tournament_data['standings'],
        pairings=tournament_data['pairings'],
        last_update=tournament_data['last_update']
    )

@app.route('/static/<path:filename>')
def serve_static(filename):
    """Serve static files like the poster image."""
    from flask import send_from_directory
    return send_from_directory('.', filename)

def start_file_watcher():
    """Start watching the Excel file for changes."""
    event_handler = ExcelFileHandler("tournament.xlsx")
    observer = Observer()
    observer.schedule(event_handler, path='.', recursive=False)
    observer.start()
    print("📂 File watcher started - monitoring tournament.xlsx for changes")
    return observer

def main():
    """Main function to start the server."""
    print("=" * 70)
    print("🏆 SWISS TOURNAMENT LIVE DISPLAY SERVER")
    print("=" * 70)
    print()
    
    # Check if Excel file exists
    if not Path("tournament.xlsx").exists():
        print("❌ ERROR: tournament.xlsx not found!")
        print("Please make sure the Excel file is in the same directory.")
        return
    
    # Initial data load
    print("📊 Loading initial tournament data...")
    load_tournament_data()
    
    # Start file watcher in background
    observer = start_file_watcher()
    
    print()
    print("🚀 Server starting...")
    print("=" * 70)
    print()
    print("🌐 Open your browser and go to:")
    print()
    print("    http://localhost:5000")
    print()
    print("=" * 70)
    print()
    print("💡 The display will auto-refresh every 10 seconds")
    print("📂 File watcher detects changes to tournament.xlsx")
    print("⚠️  Press CTRL+C to stop the server")
    print()
    
    try:
        app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)
    except KeyboardInterrupt:
        print("\n\n🛑 Stopping server...")
        observer.stop()
        observer.join()
        print("✅ Server stopped successfully")

if __name__ == "__main__":
    main()