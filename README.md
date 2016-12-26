# O/U Line - Data Testing #


## Summary ##

This is a program I made to test a variation of the [following](http://www.betandskill.com/both-teams-to-score-betting-system-btts-strategy.html) football betting strategy, by using about 70.000 football game data entries from [football-data.co.uk](http://www.football-data.co.uk/downloadm.php).

Parts of the scrips can be used to manipulate the data for various different betting strategies.

The original system goes like this:

```
Looking at the last 6 matches, we apply our own marking system 

HOME TEAM POINTS
R1) +1 point for every match when both teams scored
R2) +1 point for every home match when both teams scored
R3) +0.5 point for each goal over 2 scored per game
R4) +0.5 point for each goal over 2 conceded per game
R5) -2 points for every 0-0 draw

AWAY TEAM POINTS
R6) +1 point for every match when both teams scored
R7) +1 point for every away match when both teams scored
R8) +0.5 point for each goal over 2 scored per game
R9) +0.5 point for each goal over 2 conceded per game
R10) -2 points for every 0-0 draw


MATCH POINTS
Take into account the following rules considering the current match.

R11) +2 points if the away team is slightly favourite
R12) +1 point if the home team is slightly favourite
R13) -2 points if the home team is heavy favourite
R14) -1 point if the away team is heavy favourite

Now sum home team, away team and match points and **select only matches with total points higher than 18**.
```

## How to use ##

1. Run *xlsx_preparation.py* (change filenames to yours). The script removes all unwanted columns and converts .xls to .xlsx. The new file has the -FIXED suffix.
2. Run *point_computation.py* (change filenames to yours). It loads all FIXED .xlsx files and does the following:
	1. Calculates Team History - the last 6 game history of each team at each point in time.
	2. Calculates the **Points** (see above) for each game.
	3. Calculates the Over/Under Line for each game.
	4. Calculates the Over/Under Line odds for each game.
	5. Merges everthing (except games with no **Points**) in a single .xlsx file named UBER_SHEET.

## Changelog ##

### Version 1.1 ###

- Accessing multiple (all) .xlsx files on a single run
- Added line & line odds calculation within *point_computtion.py*
- Merges all workbooks and worksheets in a single .xlsx file for statistical ease of use


### Version 1.0 ###

* Merged two scripts in one (history discovery & point calculation) to *point_computation.py*
* Accessing multiple (all) worksheets in a single .xlsx
* Points marked BLUE if 1X2 odds are missing (+-2 points error)
* **Wrote extra script for calculating the precentage Over 2.5 games in terms of Points (e.g. 23pts - 62%)**


### Version 0.9 ###

* Accessing single worksheet
* Calculating team's history (last_6_game_data.py)
* Calculating points (point_calculator.py)

