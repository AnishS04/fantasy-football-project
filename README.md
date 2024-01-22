# Fantasy Football Dashboard 

## Overview 
I had created this project to display the statistics of the top 25 fantasy football players for the past half-decade. The data was retrieved, cleaned, and visualized to help analyze the top players and help explain common occurences within the top scorers each year. This can allow for better drafts for the future years as well as a general knowledge of what to look out for future fantasy football phenomenons. It also helps directly compare what common stars are on the incline and decline for future fantasy football seasons. 

## Data sources
https://fantasy.nfl.com/research/scoringleaders?offset=1&position=O&sort=pts&statCategory=stats&statSeason=2022&statType=seasonStats&statWeek=18#researchScoringLeaders=researchScoringLeaders%2C%2Fresearch%2Fscoringleaders%253Fposition%253DO%2526sort%253Dpts%2526statCategory%253Dstats%2526statSeason%253D2023%2526statType%253DseasonStats%2Creplace

## Data-processing 
First I parsed through the data and removed the unneccesary columns (in reverse order) because I was using indexes to delete the columns. When a column was deleted the index shift messed up the outcome of the columns that I wanted deleted. I also redid the names of the columns to better display what the purpose of the column was. The ranking that was given from the initial collection of data was improper and based on total fantasy points which could be skewed based on how many games were played, so rather I created a new rank that would be based on Fantasy Points Per Game scored, which was a better indiciatior of weekly perfomance. The tie breaker if the PPG were the same, ended up being the total fantasy points scored. New columns for team and position was extrapolated from the string column "Player" to help with queries that would be run in Power BI. Then updated the changes directly to the excel workbook. 

## Results
The results were displayed in a visualization (dashboard) using Power BI.

## Dependencies 
Pandas == 2.1.4
Openpyxl == 3.1.2
