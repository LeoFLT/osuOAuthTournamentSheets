# osu! 1v1 Tournament OAuth sheets

This collection of scripts will:

* Authorize a player via osu! and Discord OAuth2 services, meaning identity forging is impossible.
* Add the player to a spreadsheet, along with relevant information (Timestamp, user ID, Username, Rank, Avatar, etc.)
* Make players join your desired Discord server, assignining as many roles as you wish (it will just give for players already in the server).

    Please note that the sign-up process is split into two parts, the first one being the osu! authorization, where the player registers and appears in the sheet, and the Discord authorization, where they are asked to join your server and are given their roles/have their Discord Tags recorded alongside their player information on the sheet.

# Dependencies
* clasp (to push the code to a Google Apps Script project)

WIP, but works if you're willing to put in some effort (reading the instructions).\
This repo is intended as a complement to the [template spreadsheet](https://docs.google.com/spreadsheets/d/18T20x1Vaj-WTnWagOEdLIooY2qS_wB8aWUziK7zlUT0/edit), use it instead.