import flask
import os
from flask import send_from_directory
from openpyxl import load_workbook
import requests

app = flask.Flask(__name__)
app.secret_key = 'ItIsASecret'
app.debug = True

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path, 'static'), 'favicon.ico', mimetype='image/favicon.png')

@app.route('/')
@app.route('/home')
def home():
    # Load the Excel spreadsheet
    workbook = load_workbook('R:\Sports\Fantasy Hockey 2023\PlayoffPool2023.xlsx')
    sheet = workbook['shtTeamInfo']
    team_table = sheet['tblTeam']
    goalie_table = sheet['tblGoalies']

    # Retrieve the necessary columns from tblTeam
    player_column = team_table['Player']
    team_column = team_table['Team']
    goals_column = team_table['Goals']
    assists_column = team_table['Assists']
    points_before_acquiring_column = team_table['Points Before Acquiring']
    points_gordie_howe_column = team_table['Points for Gordie Howe Hattricks']
    points_conn_smythe_column = team_table['Points for Conn Smythe']

    # Retrieve the necessary columns from tblGoalies
    goalie_player_column = goalie_table['Player']
    goalie_team_column = goalie_table['Team']
    goalie_assists_column = goalie_table['Assists']
    goalie_wins_column = goalie_table['Wins']
    goalie_shutouts_column = goalie_table['Shutouts']
    goalie_points_before_acquiring_column = goalie_table['Points Before Acquiring']

    # Iterate over the rows and calculate the points for each team
    team_points = {}
    for i in range(1, len(player_column)):
        player = player_column[i].value
        team = team_column[i].value
        player_id = team_table['Player ID'][i].value
        goals = goals_column[i].value
        assists = assists_column[i].value
        points_before_acquiring = points_before_acquiring_column[i].value
        points_gordie_howe = points_gordie_howe_column[i].value
        points_conn_smythe = points_conn_smythe_column[i].value

        # Perform API calls to fetch the current goals and assists for the player
        player_stats = get_player_stats(player_id)
        if player_stats is not None:
            goals = player_stats['goals']
            assists = player_stats['assists']

        # Calculate the total points for the player
        total_points = goals + assists - points_before_acquiring + points_gordie_howe + points_conn_smythe

        # Add the points to the respective team
        if team in team_points:
            team_points[team] += total_points
        else:
            team_points[team] = total_points

    # Perform a similar iteration for the goalie table to calculate goalie points

    # Now you have the team_points dictionary with the calculated points for each team
    # You can use this data to display a table or any other format you desire

    return flask.render_template('home.html', team_points=team_points)

def get_player_stats(player_id):
    str_year = '20222023'  # Set the desired year here
    url = f'https://statsapi.web.nhl.com/api/v1/people/{player_id}/stats?stats=statsSingleSeasonPlayoffs&season={str_year}'
    response = requests.get(url)
    if response.status_code