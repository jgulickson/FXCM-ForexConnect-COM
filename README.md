# FX-ForexConnect-COM

## Overview
#### Summary
Repository contains an Excel (\*.xlsm) spreadsheet which sources select account related information from a real or demo [FXCM](https://www.fxcm.com) trading account via the [ForexConnect API](https://www.fxcm.com/uk/trading-services/api-trading/technology/) with [COM wrapper](http://fxcodebase.com/wiki/index.php/Using_ForexConnect_in_COM).  Additional information is then calculated within Excel to reach an end objective of an account dashboard.  Project was originally created as a proof of concept in 2013.

![Account-Dashboard](/README-Images/Account-Dashboard.png)

#### Note
The end deliverable is the Excel spreadsheet in this repository however all VBA code has been exported to (\*.bas) and (\*.cls) files in an attempt to improve consumability online.

#### Requirements
1. ForexConnect API ([Download](http://fxcodebase.com/wiki/index.php/Download)).

2. ForexConnect API COM  Wrapper ([Download](http://fxcodebase.com/wiki/index.php/Download)).

3. The account must be enabled for ForexConnect API access; contact FXCM for more information.

## **Installation**
1. Clone or download 'Account-Dashboard.xlsm' file from this repository.

2. Open 'Account-Dashboard.xlsm' after completed the steps outlined in 'Requirements' section above.

3. When asked, select 'Enable' macros.

## Version History
#### Account Dashboard
###### 1.x.071313
- ***Feature release***
- Updated code for optimization purposes

###### 1.x.070713
- ***Feature release***
- Added amount and size in usd to account section
- Added counter currency to open and closed trades sections
- Added currency section
- Added voice for open and closed trades

###### 1.x.070613
- ***Feature release***
- Added duration open to open trades and closed trades sections

###### 1.x.070513
- ***Feature release***
- Fixed pip cost calculation for trades where pl = 0
- Updated code for optimization purposes
- Added extensive code commenting

###### 1.x.070413
- ***Feature release***
- Added total row to accounts, open trades and closed trades sections
- Added color logic and arrow logic to column headings when sorting
- Added alternating color background for each row
- Updated leverage calculation to use size in USD for accuracy; previously was used margin
- Fixed sorting to work on trades opened and closed after sort is requested 

###### 1.x.070313
- ***Feature release***
- Added symbol, base currency, size in usd to open trades and closed trades sections
- Added sorting to accounts, open trades and closed trades sections

###### 1.x.070213
- ***Feature release***
- Added color logic for accounts, open trades and closed trades sections
- Updated GUI for usability

###### 1.x.062013
- ***Feature release***
- Created open trades and closed trades section using table manager 
- Updated GUI for usability 

###### 1.x.052713
- ***Feature release***
- Created accounts section using table manager 