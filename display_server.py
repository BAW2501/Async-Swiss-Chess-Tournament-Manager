"""
Swiss Tournament Live Display Server
Watches tournament.xlsx and displays live standings and pairings
Built by BAW2501 - https://github.com/BAW2501
"""
from markupsafe import Markup
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

# Shared CSS styles - OPTIMIZED FOR TV / HIGH DENSITY DISPLAY
SHARED_STYLES = """
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }

    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: linear-gradient(135deg, #D4A574 0%, #8B6F47 50%, #D4A574 100%);
        min-height: 100vh;
        padding: 5px; /* Reduced body padding */
        color: #2c1810;
        overflow: hidden; /* Prevent double scrollbars on TV */
    }

    .container {
        max-width: 99vw; /* Full width for TV */
        margin: 0 auto;
        height: 100vh;
        display: flex;
        flex-direction: column;
    }

    /* COMPACT HEADER */
    .header {
        text-align: center;
        margin-bottom: 10px;
        background: rgba(255, 255, 255, 0.95);
        padding: 8px 15px;
        border-radius: 10px;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.3);
        border: 2px solid #5d4e37;
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-shrink: 0;
    }

    .header-content {
        display: flex;
        align-items: center;
        gap: 20px;
        flex-grow: 1;
        justify-content: center;
    }

    .logo-container {
        margin: 0;
    }

    .logo-container img {
        height: 50px; /* Force small logo */
        width: auto;
        border-radius: 5px;
    }

    h1 {
        font-size: 1.6em; /* Smaller font */
        color: #5d4e37;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .subtitle {
        font-size: 1em;
        color: #8B6F47;
        font-weight: 600;
        margin-left: 15px;
        display: inline-block;
    }

    .last-update {
        font-size: 0.85em;
        color: #666;
        margin-left: auto;
    }

    /* COMPACT NAV */
    .nav-buttons {
        display: flex;
        gap: 8px;
        justify-content: center;
        margin-bottom: 10px;
        flex-shrink: 0;
    }

    .nav-button {
        background: rgba(255, 255, 255, 0.9);
        border: 2px solid #5d4e37;
        padding: 6px 20px; /* Smaller padding */
        font-size: 1.1em;
        font-weight: bold;
        color: #5d4e37;
        text-decoration: none;
        border-radius: 8px;
        transition: all 0.3s ease;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .nav-button:hover {
        transform: translateY(-2px);
    }

    .nav-button.active {
        background: linear-gradient(135deg, #5d4e37, #8B6F47);
        color: white;
        border-color: #8B6F47;
    }

    .filter-container {
        display: flex;
        gap: 8px;
        justify-content: center;
        margin-bottom: 8px;
        flex-shrink: 0;
    }

    .filter-button {
        background: rgba(255, 255, 255, 0.9);
        border: 1px solid #5d4e37;
        padding: 4px 15px;
        font-size: 0.9em;
        font-weight: 600;
        color: #5d4e37;
        cursor: pointer;
        border-radius: 6px;
    }

    .filter-button.active {
        background: linear-gradient(135deg, #D4A574, #8B6F47);
        color: white;
    }

    .section {
        background: rgba(255, 255, 255, 0.95);
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        border: 2px solid #5d4e37;
        display: flex;
        flex-direction: column;
        height: 100%; /* Take remaining height */
        overflow: hidden;
    }

    .section-title {
        font-size: 1.3em;
        color: #5d4e37;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
        border-bottom: 2px solid #D4A574;
        padding-bottom: 4px;
        flex-shrink: 0;
    }

    /* OPTIMIZED TABLE FOR MORE ROWS */
    .table-container {
        flex-grow: 1;
        overflow-y: auto;
        border-radius: 5px;
    }

    table {
        width: 100%;
        border-collapse: collapse;
        background: white;
    }

    thead {
        position: sticky;
        top: 0;
        z-index: 10;
        background: linear-gradient(135deg, #5d4e37 0%, #8B6F47 100%);
        color: white;
    }

    th {
        padding: 6px 8px; /* Compact headers */
        text-align: left;
        font-weight: 700;
        font-size: 0.9em;
        text-transform: uppercase;
    }

    th i { margin-right: 4px; }

    td {
        padding: 4px 8px; /* VERY COMPACT ROWS */
        border-bottom: 1px solid #e0e0e0;
        font-size: 1em; /* Keep readable */
        line-height: 1.2;
    }

    /* Zebra striping - alternating rows */
    tbody tr:nth-child(even) {
        background-color: #f7f3ed; /* Subtle beige matching your theme */
    }

    /* Hover effect - slightly darker to be visible on both row colors */
    tbody tr:hover { 
        background-color: #e8ded4; 
        transition: background-color 0.1s ease;
    }

    .rank { font-weight: bold; color: #5d4e37; }
    .rank-1 { color: #FFD700; }
    .rank-2 { color: #C0C0C0; }
    .rank-3 { color: #CD7F32; }

    .points { font-weight: bold; color: #2c5f2d; }

    .board-number {
        background: linear-gradient(135deg, #5d4e37, #8B6F47);
        color: white;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
        display: inline-block;
        min-width: 40px;
        text-align: center;
        font-size: 0.9em;
    }

    .player-cell { font-weight: 600; color: #2c1810; }

    .round-badge {
        background: #D4A574;
        color: #2c1810;
        padding: 2px 6px;
        border-radius: 10px;
        font-size: 0.8em;
        font-weight: bold;
    }

    .result {
        font-weight: bold;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 0.9em;
    }

    .result-win { background: #d4edda; color: #155724; }
    .result-draw { background: #fff3cd; color: #856404; }
    .result-loss { background: #f8d7da; color: #721c24; }
    .result-pending { background: #e2e3e5; color: #383d41; }

    /* Hide footer on TV to save space */
    .footer {
        display: none; 
    }

    .empty-state { text-align: center; padding: 20px; color: #666; }
    .empty-state i { font-size: 2em; color: #D4A574; margin-bottom: 10px; }

    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }

    .live-indicator {
        display: inline-block;
        width: 10px;
        height: 10px;
        background: #ff4444;
        border-radius: 50%;
        animation: pulse 2s infinite;
        margin-right: 5px;
    }

    .pairing-row { display: table-row; }
    .pairing-row.hidden { display: none; }
"""

# Shared header component
def render_header(last_update):
    return f"""
    <div class="header">
        <div class="header-content">
            <div class="logo-container">
                <img src="/static/tournament_poster.jpg" alt="Logo" onerror="this.style.display='none'">
            </div>
            <h1>
                <i class="fas fa-chess"></i>
                Amateur Chess Tournament
            </h1>
            <div class="subtitle">
                <span class="live-indicator"></span>
                LIVE
            </div>
        </div>
        {'<div class="last-update"><i class="far fa-clock"></i> ' + last_update + '</div>' if last_update else ''}
    </div>
    """

# Navigation component
def render_nav(active_page):
    pages = [
        ('/', 'combined', 'fas fa-columns', 'Combined'),
        ('/standings', 'standings', 'fas fa-ranking-star', 'Standings'),
        ('/pairings', 'pairings', 'fas fa-chess-board', 'Pairings')
    ]
    
    nav_html = '<div class="nav-buttons">'
    for url, page_id, icon, label in pages:
        active_class = 'active' if page_id == active_page else ''
        nav_html += f'<a href="{url}" class="nav-button {active_class}"><i class="{icon}"></i>{label}</a>'
    nav_html += '</div>'
    return nav_html

# Standings table component
def render_standings_table(standings):
    if not standings:
        return '<div class="empty-state"><i class="fas fa-hourglass-start"></i><p>Waiting for standings...</p></div>'
    
    html = '<div class="table-container"><table><thead><tr>'
    html += '<th>Rank</th>'
    html += '<th>Player</th>'
    html += '<th>Pts</th>'
    html += '<th>BucT</th>'
    html += '<th>DE</th>'
    html += '<th>Ber</th>'
    html += '</tr></thead><tbody>'
    
    for player in standings:
        rank_class = ''
        icon = ''
        if player['Pos'] == 1:
            rank_class = 'rank-1'
            icon = '<i class="fas fa-crown"></i> '
        elif player['Pos'] == 2:
            rank_class = 'rank-2'
            icon = '<i class="fas fa-medal"></i> '
        elif player['Pos'] == 3:
            rank_class = 'rank-3'
            icon = '<i class="fas fa-award"></i> '
        
        html += f'<tr><td class="rank {rank_class}">{icon}{player["Pos"]}</td>'
        html += f'<td class="player-cell">{player["Player Name"]}</td>'
        html += f'<td class="points">{player["Pt"]}</td>'
        html += f'<td>{player["BucT"]}</td>'
        html += f'<td>{player["DE"]}</td>'
        html += f'<td>{player["Ber"]}</td></tr>'
    
    html += '</tbody></table></div>'
    return html

# Pairings table component
def render_pairings_table(pairings, show_filter=True):
    if not pairings:
        return '<div class="empty-state"><i class="fas fa-chess-knight"></i><p>Waiting for pairings...</p></div>'
    
    html = ''
    if show_filter:
        html += """
        <div class="filter-container">
            <button class="filter-button active" onclick="filterPairings('all')">All</button>
            <button class="filter-button" onclick="filterPairings('pending')">Live</button>
            <button class="filter-button" onclick="filterPairings('finished')">Done</button>
        </div>
        """
    
    html += '<div class="table-container"><table><thead><tr>'
    html += '<th>Round</th>'
    html += '<th>Board</th>'
    html += '<th>White</th>'
    html += '<th></th>'
    html += '<th>Black</th>'
    html += '<th>Result</th>'
    html += '</tr></thead><tbody id="pairings-tbody">'
    
    for pairing in pairings:
        status = pairing['result_status']
        html += f'<tr class="pairing-row" data-status="{status}">'
        html += f'<td><span class="round-badge">{pairing["Round"]}</span></td>'
        html += f'<td><span class="board-number">{pairing["Board"]}</span></td>'
        html += f'<td class="player-cell">{pairing["White Name"]}</td>'
        html += '<td class="vs-separator" style="font-size:0.8em; color:#ccc;">vs</td>'
        html += f'<td class="player-cell">{pairing["Black Name"]}</td>'
        
        result_html = '<span class="result result-pending">Live</span>'
        if status == 'White Win':
            result_html = '<span class="result result-win">1-0</span>'
        elif status == 'Black Win':
            result_html = '<span class="result result-loss">0-1</span>'
        elif status == 'Draw':
            result_html = '<span class="result result-draw">½-½</span>'
        
        html += f'<td>{result_html}</td></tr>'
    
    html += '</tbody></table></div>'
    return html

# Footer component (Hidden via CSS, but function kept for structure)
def render_footer():
    return """
    <div class="footer">
        <p>Built by BAW2501</p>
    </div>
    """

# Filter script
FILTER_SCRIPT = """
<script>
function filterPairings(filterType) {
    const rows = document.querySelectorAll('.pairing-row');
    const buttons = document.querySelectorAll('.filter-button');
    
    buttons.forEach(btn => btn.classList.remove('active'));
    event.target.closest('.filter-button').classList.add('active');
    
    rows.forEach(row => {
        const status = row.getAttribute('data-status');
        if (filterType === 'all') {
            row.classList.remove('hidden');
        } else if (filterType === 'pending') {
            row.classList.toggle('hidden', status !== 'Pending');
        } else if (filterType === 'finished') {
            row.classList.toggle('hidden', status === 'Pending');
        }
    });
}
</script>
"""

# Combined view page
COMBINED_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tournament Display</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        {{ styles }}
        .split-view {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
            flex-grow: 1;
            overflow: hidden;
            padding-bottom: 5px;
        }
        @media (max-width: 1200px) {
            .split-view {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        {{ header }}
        {{ nav }}
        <div class="split-view">
            <div class="section">
                <h2 class="section-title"><i class="fas fa-ranking-star"></i> Standings</h2>
                {{ standings_table }}
            </div>
            <div class="section">
                <h2 class="section-title"><i class="fas fa-chess-board"></i> Pairings</h2>
                {{ pairings_table }}
            </div>
        </div>
        {{ footer }}
    </div>
    {{ filter_script }}
</body>
</html>
"""

# Single view page template
SINGLE_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        {{ styles }}
        .single-view-container {
            flex-grow: 1;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            padding-bottom: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        {{ header }}
        {{ nav }}
        <div class="single-view-container">
            <div class="section">
                <h2 class="section-title">{{ section_title }}</h2>
                {{ content }}
            </div>
        </div>
        {{ footer }}
    </div>
    {{ extra_script }}
</body>
</html>
"""

class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, filepath):
        self.filepath = filepath
        self.last_modified = 0
        
    def on_modified(self, event):
        if event.src_path.endswith('tournament.xlsx'):
            current_time = time.time()
            if current_time - self.last_modified > 2:
                self.last_modified = current_time
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Excel file changed - reloading data...")
                time.sleep(0.5)
                load_tournament_data()

def load_tournament_data():
    try:
        filepath = "tournament.xlsx"
        
        try:
            standings_df = pd.read_excel(filepath, sheet_name="Standings", engine='openpyxl')
            tournament_data['standings'] = standings_df.to_dict('records')
        except Exception as e:
            print(f"Could not load standings: {e}")
            tournament_data['standings'] = []
        
        try:
            pairings_df = pd.read_excel(filepath, sheet_name="Pairings", engine='openpyxl')
            pairings_list = []
            for _, row in pairings_df.iterrows():
                pairing = row.to_dict()
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
        
        tournament_data['last_update'] = datetime.now().strftime('%H:%M')
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Data loaded: {len(tournament_data['standings'])} standings, {len(tournament_data['pairings'])} pairings")
    except Exception as e:
        print(f"Error loading tournament data: {e}")

@app.route('/')
def combined_view():
    return render_template_string(
        COMBINED_PAGE,
        styles=Markup(SHARED_STYLES),
        header=Markup(render_header(tournament_data['last_update'])),
        nav=Markup(render_nav('combined')),
        standings_table=Markup(render_standings_table(tournament_data['standings'])),
        pairings_table=Markup(render_pairings_table(tournament_data['pairings'], show_filter=True)),
        footer=Markup(render_footer()),
        filter_script=Markup(FILTER_SCRIPT)
    )

@app.route('/standings')
def standings_view():
    return render_template_string(
        SINGLE_PAGE,
        title='Standings',
        styles=Markup(SHARED_STYLES),
        header=Markup(render_header(tournament_data['last_update'])),
        nav=Markup(render_nav('standings')),
        section_title=Markup('<i class="fas fa-ranking-star"></i> Standings'),
        content=Markup(render_standings_table(tournament_data['standings'])),
        footer=Markup(render_footer()),
        extra_script=''
    )

@app.route('/pairings')
def pairings_view():
    return render_template_string(
        SINGLE_PAGE,
        title='Pairings',
        styles=Markup(SHARED_STYLES),
        header=Markup(render_header(tournament_data['last_update'])),
        nav=Markup(render_nav('pairings')),
        section_title=Markup('<i class="fas fa-chess-board"></i> Pairings'),
        content=Markup(render_pairings_table(tournament_data['pairings'], show_filter=True)),
        footer=Markup(render_footer()),
        extra_script=Markup(FILTER_SCRIPT)
    )

@app.route('/static/<path:filename>')
def serve_static(filename):
    from flask import send_from_directory
    return send_from_directory('.', filename)

def start_file_watcher():
    event_handler = ExcelFileHandler("tournament.xlsx")
    observer = Observer()
    observer.schedule(event_handler, path='.', recursive=False)
    observer.start()
    print("📂 File watcher started - monitoring tournament.xlsx for changes")
    return observer

def main():
    print("=" * 70)
    print("🏆 SWISS TOURNAMENT LIVE DISPLAY SERVER")
    print("=" * 70)
    print()
    
    if not Path("tournament.xlsx").exists():
        print("❌ ERROR: tournament.xlsx not found!")
        print("Please make sure the Excel file is in the same directory.")
        return
    
    print("📊 Loading initial tournament data...")
    load_tournament_data()
    
    observer = start_file_watcher()
    
    print()
    print("🚀 Server starting...")
    print("=" * 70)
    print()
    print("🌐 Open your browser:")
    print()
    print("    http://localhost:5000           - Combined View (both tables)")
    print("    http://localhost:5000/standings - Standings Only")
    print("    http://localhost:5000/pairings  - Pairings Only")
    print()
    print("=" * 70)
    print()
    print("📂 File watcher detects changes to tournament.xlsx")
    print("🔄 Manually refresh browser to see updates")
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