-- Tournament Planner .TP schema (Visual Reality BTP 2026.3)
-- Reverse-engineered from: NW Doppelturnier 26-27.TP
-- Format: Microsoft Jet 4 (Access 2000). 70 tables.
-- Generated 2026-08-29 with mdb-schema.

-- ============================================================
-- FOREIGN KEYS DECLARED IN MSysRelationships (the only enforced ones)
-- ============================================================
--   ClubTeam                 Player.club -> Club.id
--   CourtTeamMatch           PlayerMatch.court -> Court.id
--   CourtTypeCourt           Court.courttype -> CourtType.id
--   DrawTeamWedstrijd        PlayerMatch.draw -> Draw.id
--   EntryTeamWedstrijd       PlayerMatch.entry -> Entry.id
--   KlasseEntry              Entry.event -> Event.id
--   KlassePoule              Draw.event -> Event.id
--   LinkTeamMatch            PlayerMatch.link -> Link.id
--   LocationCourt            Court.location -> Location.id
--   OfficialFunctionMatchOfficial MatchOfficial.function -> OfficialFunction.id
--   OfficialMatchOfficial    MatchOfficial.official -> Official.id
--   TeamEntry                Entry.player1 -> Player.id
--   TeamMatchMatchOfficial   MatchOfficial.match -> PlayerMatch.id

-- ============================================================
-- VERIFIED BY CONVENTION (not declared; confirmed against real rows)
-- ============================================================
--   Entry.player2                  -> Player.id
--   PlayerMatch.event              -> Event.id
--   PlayerMatch.winner             -> Entry.id
--   Court.playermatch              -> PlayerMatch.id
--   Payment.player                 -> Player.id
--   Player.level1 / level2         -> PlayerLevel.ID     (note: uppercase ID)
--   Player.country                 -> Country.ID         (note: uppercase ID)
--   TournamentTime.location        -> Location.id
--
--   BRACKET TREE - composite, NOT by row id:
--   PlayerMatch.van1 / van2        -> PlayerMatch.planning  WHERE same draw
--   PlayerMatch.wn                 -> PlayerMatch.planning  WHERE same draw  (winner advances to)
--   PlayerMatch.vn                 -> PlayerMatch.planning  WHERE same draw  (loser drops to)
--   (draw, planning) is unique. Joining van1 to PlayerMatch.id returns ZERO rows.
--
--   NOT a foreign key despite the name:
--   TournamentTime.TournamentDay   is a DATE value, not a reference to TournamentDay.id
--   Availability.objectid          polymorphic, discriminated by Availability.objecttype

-- ============================================================
-- TABLE DEFINITIONS
-- ============================================================
-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

-- That file uses encoding UTF-8

CREATE TABLE [Availability]
 (
	[id]			Long Integer, 
	[objecttype]			Long Integer, 
	[objectid]			Long Integer, 
	[day]			DateTime, 
	[availability]			OLE (255)
);

CREATE TABLE [Club]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[clubid]			Text (20), 
	[district]			Long Integer, 
	[contact]			Text (50), 
	[address]			Text (50), 
	[address2]			Text (50), 
	[address3]			Text (50), 
	[postalcode]			Text (10), 
	[city]			Text (50), 
	[state]			Text (50), 
	[countryid]			Long Integer, 
	[phone]			Text (20), 
	[fax]			Text (20), 
	[mobile]			Text (20), 
	[email]			Text (50), 
	[website]			Text (100), 
	[abbreviation]			Text (25), 
	[memo]			Memo/Hyperlink (255)
);

CREATE TABLE [CodeViolation]
 (
	[id]			Long Integer, 
	[foreignid]			Text (50), 
	[offencetype]			Long Integer, 
	[entryoffence]			Long Integer, 
	[match]			Long Integer, 
	[step]			Long Integer, 
	[setscore]			Text (20), 
	[gamescore]			Text (20), 
	[pointsscore]			Text (20), 
	[player]			Long Integer, 
	[code]			Long Integer, 
	[description]			Memo/Hyperlink (255), 
	[sortorder]			Long Integer, 
	[official]			Long Integer, 
	[cvdate]			DateTime, 
	[fine]			Currency, 
	[withdrawalmethod]			Long Integer, 
	[reason]			Long Integer, 
	[followup]			Boolean NOT NULL, 
	[tournamentkey]			Text (20), 
	[event]			Long Integer
);

CREATE TABLE [Country]
 (
	[id]			Long Integer, 
	[code]			Text (3) NOT NULL, 
	[name]			Text (30), 
	[image]			Long Integer
);

CREATE TABLE [CourtType]
 (
	[id]			Long Integer, 
	[name]			Text (50)
);

CREATE TABLE [District]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[code]			Text (50), 
	[abbreviation]			Text (25)
);

CREATE TABLE [Draw]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[event]			Long Integer, 
	[stage]			Long Integer, 
	[drawtype]			Long Integer NOT NULL, 
	[drawsize]			Long Integer NOT NULL, 
	[playoff]			Boolean NOT NULL, 
	[playoffsize]			Long Integer, 
	[consolation]			Long Integer, 
	[consolationsize]			Long Integer, 
	[consolationplayoffsize]			Long Integer, 
	[qualification]			Boolean NOT NULL, 
	[drawgroup]			Long Integer, 
	[drawcolumns]			Long Integer, 
	[position]			Long Integer, 
	[header]			Long Integer, 
	[footer]			Long Integer, 
	[startdate]			DateTime, 
	[enddate]			DateTime, 
	[printmodified]			Boolean NOT NULL, 
	[consolationplayoff]			Boolean NOT NULL, 
	[lastfeedinround]			Long Integer, 
	[qualificationrounds]			Long Integer, 
	[drawendsize]			Long Integer, 
	[drawrounds]			Long Integer, 
	[maxentries]			Long Integer, 
	[topacc]			Text (25), 
	[cutoff]			Text (25), 
	[lastdirectacceptance]			Text (100), 
	[representative1]			Long Integer, 
	[representative2]			Long Integer, 
	[drawdate]			DateTime, 
	[topseed]			Text (25), 
	[lastseed]			Text (25), 
	[revised]			Boolean NOT NULL, 
	[location]			Long Integer, 
	[court]			Long Integer, 
	[showonttv]			Boolean NOT NULL, 
	[fixture]			Long Integer, 
	[fixturetemplate]			Long Integer, 
	[grading]			Long Integer
);

CREATE TABLE [drawformat]
 (
	[id]			Long Integer, 
	[name]			Text (100), 
	[isdefault]			Boolean NOT NULL
);

CREATE TABLE [Entry]
 (
	[id]			Long Integer, 
	[event]			Long Integer NOT NULL, 
	[player1]			Long Integer NOT NULL, 
	[player2]			Long Integer, 
	[seed1]			Long Integer, 
	[seed2]			Long Integer, 
	[status]			Long Integer, 
	[partnerwanted]			Boolean NOT NULL, 
	[exclude]			Boolean NOT NULL, 
	[qseed1]			Long Integer, 
	[qseed2]			Long Integer, 
	[qstatus]			Long Integer, 
	[entrytype]			Long Integer, 
	[entrylist]			Long Integer, 
	[entered]			Boolean NOT NULL, 
	[accepttb]			Long Integer, 
	[seedtb]			Long Integer, 
	[signedin]			Long Integer, 
	[signedinconsolation]			Long Integer, 
	[accepttype]			Long Integer, 
	[acceptposition]			Long Integer, 
	[protectedranking]			Boolean NOT NULL, 
	[itforder]			Long Integer, 
	[protectedranking2]			Boolean NOT NULL, 
	[chip]			Long Integer, 
	[qchip]			Long Integer, 
	[qseedtb]			Long Integer, 
	[cseed1]			Long Integer, 
	[cseed2]			Long Integer, 
	[cseedtb]			Long Integer, 
	[lltb]			Long Integer, 
	[reserve]			Long Integer, 
	[alternatetb]			Long Integer, 
	[protectedrankingD]			Boolean NOT NULL, 
	[protectedranking2D]			Boolean NOT NULL, 
	[withdrawdate]			DateTime, 
	[withdrawreason]			Text (80), 
	[playup]			Boolean NOT NULL, 
	[recentform]			Long Integer, 
	[onlinepartner]			Text (80), 
	[validationstatus]			Long Integer, 
	[validationmessage]			Text (255)
);

CREATE TABLE [Entryformitem]
 (
	[id]			Long Integer, 
	[fieldname]			Text (80), 
	[fieldtype]			Long Integer, 
	[displayorder]			Long Integer, 
	[fee]			Currency, 
	[mandatory]			Boolean NOT NULL
);

CREATE TABLE [Event]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[abbreviation]			Text (10), 
	[gender]			Long Integer, 
	[eventtype]			Long Integer, 
	[minlevel]			Long Integer, 
	[level]			Long Integer, 
	[min_age]			Long Integer, 
	[max_age]			Long Integer, 
	[fee]			Currency, 
	[separateseeding]			Boolean NOT NULL, 
	[allowonlineentry]			Boolean NOT NULL, 
	[drawcomposition]			Long Integer, 
	[qualcomposition]			Long Integer, 
	[maindrawsize]			Long Integer, 
	[mainmaxentries]			Long Integer, 
	[qualdrawsize]			Long Integer, 
	[qualendsize]			Long Integer, 
	[qualmaxentries]			Long Integer, 
	[consolation]			Long Integer, 
	[foreignid]			Text (255), 
	[tournamentid]			Long Integer, 
	[scoringformat]			Long Integer, 
	[scoringformatmain]			Long Integer, 
	[scoringformatqual]			Long Integer, 
	[scoringformatcons]			Long Integer, 
	[grading]			Long Integer, 
	[subgrading]			Long Integer, 
	[subgrading2]			Long Integer, 
	[parentevent]			Long Integer, 
	[numranked]			Long Integer, 
	[points1]			Double, 
	[points2]			Double, 
	[maxonlineentries]			Long Integer, 
	[maxreserveentries]			Long Integer, 
	[min_weight]			Long Integer, 
	[max_weight]			Long Integer, 
	[weighstart]			DateTime, 
	[weighfinish]			DateTime, 
	[starttime]			DateTime, 
	[matchduration]			Long Integer, 
	[matchbreak]			Long Integer, 
	[block]			Long Integer, 
	[league]			Long Integer, 
	[paraclass]			Long Integer, 
	[location]			Long Integer, 
	[tournamentinformationid]			Long Integer, 
	[drawformat]			Long Integer, 
	[startdate]			DateTime, 
	[enddate]			DateTime, 
	[entrystartdate]			DateTime, 
	[entryenddate]			DateTime, 
	[withdrawaldate]			DateTime, 
	[ltaentryrestrictions]			Long Integer
);

CREATE TABLE [eventpart]
 (
	[ID]			Long Integer, 
	[event]			Long Integer, 
	[block]			Long Integer, 
	[eventscheduleitem]			Long Integer, 
	[partnr]			Long Integer, 
	[totalparts]			Long Integer, 
	[duration]			Long Integer, 
	[actualduration]			Long Integer
);

CREATE TABLE [eventprizemoney]
 (
	[id]			Long Integer, 
	[event]			Long Integer, 
	[position1]			Long Integer, 
	[position2]			Long Integer, 
	[amount]			Currency
);

CREATE TABLE [eventschedule]
 (
	[id]			Long Integer, 
	[day]			Long Integer, 
	[numcourts]			Long Integer
);

CREATE TABLE [eventscheduleitem]
 (
	[ID]			Long Integer, 
	[block]			Long Integer, 
	[court]			Long Integer, 
	[starttime]			DateTime, 
	[endtime]			DateTime
);

CREATE TABLE [files]
 (
	[ID]			Long Integer, 
	[filedata]			OLE (255), 
	[filename]			Text (255), 
	[filetype]			Long Integer, 
	[filesize]			Long Integer, 
	[filedate]			DateTime, 
	[tag]			Long Integer
);

CREATE TABLE [fixture]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[fixturetemplate]			Long Integer
);

CREATE TABLE [fixtureitem]
 (
	[id]			Long Integer, 
	[round]			Long Integer, 
	[playtime]			DateTime, 
	[fixture]			Long Integer
);

CREATE TABLE [fixturetemplate]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[drawsize]			Long Integer, 
	[stages]			Long Integer, 
	[isdefault]			Boolean NOT NULL
);

CREATE TABLE [fixturetemplateitem]
 (
	[id]			Long Integer, 
	[team1]			Long Integer, 
	[team2]			Long Integer, 
	[round]			Long Integer, 
	[stage]			Long Integer, 
	[rownr]			Long Integer, 
	[matchnr]			Long Integer, 
	[location]			Long Integer, 
	[fixturetemplate]			Long Integer
);

CREATE TABLE [Income]
 (
	[id]			Long Integer, 
	[name]			Text (80), 
	[amount]			Currency, 
	[incometype]			Long Integer, 
	[sortorder]			Long Integer
);

CREATE TABLE [Link]
 (
	[id]			Long Integer, 
	[src_draw]			Long Integer, 
	[src_pos]			Long Integer, 
	[intsrc_pos]			Long Integer, 
	[name]			Text (50)
);

CREATE TABLE [Location]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[address]			Text (50), 
	[postalcode]			Text (10), 
	[city]			Text (50), 
	[state]			Text (50), 
	[phone]			Text (20), 
	[fax]			Text (20), 
	[countryid]			Long Integer
);

CREATE TABLE [Log]
 (
	[id]			Long Integer, 
	[actionid]			Long Integer, 
	[actiongroupid]			Long Integer, 
	[timestamp]			DateTime, 
	[description]			Memo/Hyperlink (255), 
	[data1]			Long Integer
);

CREATE TABLE [mailtemplate]
 (
	[ID]			Long Integer, 
	[name]			Text (80), 
	[template]			Memo/Hyperlink (255)
);

CREATE TABLE [MatchOfficial]
 (
	[match]			Long Integer, 
	[official]			Long Integer, 
	[function]			Long Integer
);

CREATE TABLE [Message]
 (
	[id]			Long Integer, 
	[player]			Long Integer, 
	[messagetype]			Long Integer, 
	[emailtype]			Long Integer, 
	[created]			DateTime, 
	[sent]			DateTime, 
	[status]			Long Integer, 
	[guid]			Text (50), 
	[subject]			Text (100), 
	[body]			Memo/Hyperlink (255), 
	[html]			Memo/Hyperlink (255), 
	[email]			Text (100)
);

CREATE TABLE [MessageMatch]
 (
	[id]			Long Integer, 
	[message]			Long Integer, 
	[match]			Long Integer, 
	[playtime]			DateTime
);

CREATE TABLE [Notes]
 (
	[id]			Long Integer, 
	[checked]			Boolean NOT NULL, 
	[player]			Long Integer, 
	[timestamp]			DateTime, 
	[note]			Memo/Hyperlink (255), 
	[duedate]			DateTime
);

CREATE TABLE [Official]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[address]			Text (50), 
	[phone]			Text (20), 
	[office]			Text (20), 
	[mobile]			Text (20), 
	[fax]			Text (20), 
	[postalcode]			Text (10), 
	[city]			Text (50), 
	[state]			Text (50), 
	[country]			Text (50), 
	[email]			Text (50), 
	[countryid]			Long Integer, 
	[club]			Long Integer, 
	[lastname]			Text (50), 
	[firstname]			Text (50), 
	[middlename]			Text (50), 
	[certification]			Long Integer, 
	[startdate]			DateTime, 
	[enddate]			DateTime, 
	[itffunction]			Long Integer, 
	[remarks]			Text (80), 
	[workedqualies]			Boolean NOT NULL, 
	[evaluation]			Boolean NOT NULL, 
	[chairumpirecertification]			Long Integer, 
	[chiefumpirecertification]			Long Integer, 
	[refereecertification]			Long Integer, 
	[officialid]			Text (50), 
	[asianname]			Boolean NOT NULL
);

CREATE TABLE [OfficialFunction]
 (
	[id]			Long Integer, 
	[name]			Text (50)
);

CREATE TABLE [onlineentry]
 (
	[ID]			Long Integer, 
	[player]			Long Integer, 
	[event]			Long Integer, 
	[partner]			Text (80), 
	[partnermemberID]			Text (50)
);

CREATE TABLE [OrderOfPlay]
 (
	[id]			Long Integer, 
	[day]			Long Integer, 
	[numrounds]			Long Integer, 
	[released]			DateTime, 
	[courts]			Text (255), 
	[text1]			Text (255), 
	[text2]			Text (255), 
	[text3]			Text (255), 
	[headertext]			Text (255), 
	[maxroundsperpage]			Long Integer, 
	[maxcourtsperpage]			Long Integer, 
	[landscape]			Boolean NOT NULL, 
	[headertextcolor]			Long Integer, 
	[locked]			Boolean NOT NULL
);

CREATE TABLE [OrderOfPlayCourt]
 (
	[id]			Long Integer, 
	[orderofplay]			Long Integer, 
	[court]			Long Integer, 
	[sortorder]			Long Integer
);

CREATE TABLE [OrderOfPlayItem]
 (
	[id]			Long Integer, 
	[orderofplay]			Long Integer, 
	[planning]			Long Integer, 
	[match]			Long Integer, 
	[freetext]			Text (50), 
	[time]			DateTime, 
	[orderofplaytype]			Long Integer, 
	[tofinish]			Boolean NOT NULL
);

CREATE TABLE [ParaClass]
 (
	[id]			Long Integer NOT NULL, 
	[name]			Text (20)
);

CREATE TABLE [Payment]
 (
	[id]			Long Integer, 
	[player]			Long Integer, 
	[club]			Long Integer, 
	[amount]			Currency, 
	[paymenttype]			Long Integer, 
	[paymentdate]			DateTime, 
	[code]			Text (50), 
	[payerid]			Text (50), 
	[payername]			Text (50)
);

CREATE TABLE [Playerentryformitem]
 (
	[id]			Long Integer, 
	[playerid]			Long Integer, 
	[entryformitemid]			Long Integer, 
	[quantity]			Long Integer, 
	[textvalue]			Text (255)
);

CREATE TABLE [PlayerLevel]
 (
	[id]			Long Integer NOT NULL, 
	[leveltype]			Long Integer, 
	[name]			Text (20), 
	[levelorder]			Long Integer
);

CREATE TABLE [PlayerlevelEntry]
 (
	[id]			Long Integer, 
	[leveltype]			Long Integer, 
	[playerid]			Long Integer, 
	[level1]			Long Integer, 
	[level2]			Long Integer, 
	[level3]			Long Integer
);

CREATE TABLE [PlayerMatch]
 (
	[id]			Long Integer, 
	[event]			Long Integer, 
	[draw]			Long Integer, 
	[planning]			Long Integer, 
	[entry]			Long Integer, 
	[winner]			Long Integer, 
	[link]			Long Integer, 
	[plandate]			DateTime, 
	[court]			Long Integer, 
	[location]			Long Integer, 
	[van1]			Long Integer, 
	[van2]			Long Integer, 
	[wn]			Long Integer, 
	[vn]			Long Integer, 
	[walkover]			Boolean NOT NULL, 
	[retired]			Boolean NOT NULL, 
	[team1set1]			Long Integer, 
	[team2set1]			Long Integer, 
	[team1set2]			Long Integer, 
	[team2set2]			Long Integer, 
	[team1set3]			Long Integer, 
	[team2set3]			Long Integer, 
	[team1set4]			Long Integer, 
	[team2set4]			Long Integer, 
	[team1set5]			Long Integer, 
	[team2set5]			Long Integer, 
	[team1set6]			Long Integer, 
	[team2set6]			Long Integer, 
	[team1set7]			Long Integer, 
	[team2set7]			Long Integer, 
	[status]			Long Integer, 
	[matchorder]			Long Integer, 
	[duration]			Long Integer, 
	[ranking]			Long Integer, 
	[matchno]			Long Integer, 
	[scorestatus]			Long Integer, 
	[starttime]			DateTime, 
	[highlight]			Long Integer, 
	[note]			Memo/Hyperlink (255), 
	[notetimestamp]			DateTime, 
	[scoresheetprinted]			Boolean NOT NULL, 
	[official1]			Long Integer, 
	[official2]			Long Integer, 
	[forwardloser]			Boolean NOT NULL, 
	[set1tiebreak]			Long Integer, 
	[set2tiebreak]			Long Integer, 
	[set3tiebreak]			Long Integer, 
	[set4tiebreak]			Long Integer, 
	[set5tiebreak]			Long Integer, 
	[set6tiebreak]			Long Integer, 
	[set7tiebreak]			Long Integer, 
	[shuttles]			Long Integer, 
	[prevplaytime]			DateTime, 
	[endtime]			DateTime, 
	[showbye]			Boolean NOT NULL, 
	[scoringformat]			Long Integer, 
	[scoringformatnew]			Long Integer, 
	[reversehomeaway]			Boolean NOT NULL, 
	[stage]			Long Integer, 
	[matchnr]			Long Integer, 
	[roundnr]			Long Integer, 
	[teamlocation]			Long Integer, 
	[livescore]			Long Integer, 
	[code]			Text (50), 
	[videostreaming]			Long Integer, 
	[lastupdated]			DateTime, 
	[resultoverride]			Long Integer
);

CREATE TABLE [playerprizemoney]
 (
	[id]			Long Integer, 
	[player]			Long Integer, 
	[event]			Long Integer, 
	[position1]			Long Integer, 
	[position2]			Long Integer, 
	[amount]			Currency
);

CREATE TABLE [playerutr]
 (
	[id]			Long Integer, 
	[playerid]			Long Integer, 
	[singlespoints]			Double, 
	[singlespoints2]			Double, 
	[singlesstatus]			Long Integer, 
	[singlesmasked]			Long Integer, 
	[doublespoints]			Double, 
	[doublespoints2]			Double, 
	[doublesstatus]			Long Integer, 
	[doublesmasked]			Long Integer, 
	[colorballrating]			Long Integer, 
	[colorballratingvalue]			Long Integer
);

CREATE TABLE [playerwtn]
 (
	[id]			Long Integer, 
	[playerid]			Long Integer, 
	[singlespoints]			Double, 
	[singlesconfidence]			Integer, 
	[doublespoints]			Double, 
	[doublesconfidence]			Integer
);

CREATE TABLE [PrintSettings]
 (
	[drawid]			Long Integer, 
	[sheet]			Long Integer, 
	[zoom]			Long Integer, 
	[fittopage]			Boolean NOT NULL, 
	[center]			Boolean NOT NULL, 
	[orientation]			Long Integer
);

CREATE TABLE [RankingCategory]
 (
	[ID]			Long Integer, 
	[name]			Text (50), 
	[gender]			Long Integer, 
	[gametype]			Long Integer, 
	[minage]			Long Integer, 
	[maxage]			Long Integer, 
	[level]			Long Integer, 
	[paraclass]			Long Integer, 
	[rankingtype]			Long Integer, 
	[edition]			Long Integer, 
	[rankingcountry]			Text (3)
);

CREATE TABLE [RankingEntry]
 (
	[ID]			Long Integer, 
	[rankingcategory]			Long Integer, 
	[playerid]			Long Integer, 
	[rank]			Long Integer, 
	[points]			Double, 
	[country]			Text (3)
);

CREATE TABLE [RatingSection]
 (
	[ID]			Long Integer, 
	[event]			Long Integer, 
	[rating1]			Double, 
	[rating2]			Double
);

CREATE TABLE [Replacement]
 (
	[id]			Long Integer, 
	[draw]			Long Integer, 
	[player1]			Long Integer, 
	[player2]			Long Integer, 
	[replacing1]			Long Integer, 
	[replacing2]			Long Integer
);

CREATE TABLE [seedingrule]
 (
	[ID]			Long Integer, 
	[name]			Text (50), 
	[isdefault]			Boolean NOT NULL, 
	[seedmethod]			Long Integer, 
	[seedpergroup]			Boolean NOT NULL, 
	[numseedspergroup]			Long Integer, 
	[seedpersection]			Boolean NOT NULL, 
	[numseedspersection]			Long Integer
);

CREATE TABLE [seedingruleitem]
 (
	[ID]			Long Integer, 
	[seedingrule]			Long Integer, 
	[maxentries]			Long Integer, 
	[numseeds]			Long Integer
);

CREATE TABLE [Settings]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[value]			Text (100)
);

CREATE TABLE [SettingsFloat]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[value]			Double
);

CREATE TABLE [SettingsMemo]
 (
	[id]			Long Integer, 
	[name]			Text (50) NOT NULL, 
	[value]			Memo/Hyperlink (255)
);

CREATE TABLE [stage]
 (
	[ID]			Long Integer, 
	[name]			Text (50), 
	[event]			Long Integer, 
	[displayorder]			Long Integer, 
	[stagetype]			Long Integer, 
	[drawgroup]			Long Integer, 
	[scoringformat]			Long Integer, 
	[scoringformatcons]			Long Integer
);

CREATE TABLE [TournamentDay]
 (
	[id]			Long Integer, 
	[tournamentday]			DateTime, 
	[availability]			Boolean NOT NULL, 
	[startavailability]			DateTime, 
	[endavailability]			DateTime
);

CREATE TABLE [TournamentInformation]
 (
	[ID]			Long Integer, 
	[organization]			Long Integer, 
	[tournamentid]			Text (50), 
	[informationtype]			Long Integer, 
	[tournamentname]			Text (80), 
	[applicationcode]			Text (80), 
	[itfkeytype]			Long Integer, 
	[category]			Long Integer, 
	[downloaddate]			DateTime, 
	[webcode]			Text (50), 
	[linkcode]			Text (50), 
	[unicode]			Text (50), 
	[webid]			Text (50), 
	[startdate]			DateTime, 
	[enddate]			DateTime
);

CREATE TABLE [tournamentschedule]
 (
	[ID]			Long Integer, 
	[tournamentday]			Long Integer, 
	[location]			Long Integer, 
	[match]			Long Integer, 
	[matchfilter]			Double, 
	[matchorder]			Long Integer, 
	[sortorder]			Long Integer
);

CREATE TABLE [TournamentTime]
 (
	[id]			Long Integer, 
	[tournamenttime]			DateTime, 
	[tournamentday]			DateTime, 
	[indexnr]			Long Integer, 
	[location]			Long Integer, 
	[courts]			Long Integer
);

CREATE TABLE [usagelog]
 (
	[id]			Long Integer, 
	[timestamp]			DateTime, 
	[mid]			Text (50), 
	[username]			Text (80), 
	[computername]			Text (80), 
	[license]			Text (80), 
	[productversion]			Long Integer, 
	[productsubversion]			Long Integer, 
	[rid]			Text (50)
);

CREATE TABLE [Withdrawal]
 (
	[id]			Long Integer, 
	[player]			Long Integer, 
	[event]			Long Integer, 
	[withdrawdate]			DateTime, 
	[round]			Text (50), 
	[withdrawtype]			Long Integer, 
	[reason]			Text (255), 
	[nexttournament]			Text (50), 
	[medicalcertificate]			Boolean NOT NULL, 
	[followup]			Boolean NOT NULL
);

CREATE TABLE [Court]
 (
	[id]			Long Integer, 
	[name]			Text (20), 
	[location]			Long Integer, 
	[playermatch]			Long Integer, 
	[courttype]			Long Integer, 
	[courtsurface]			Long Integer, 
	[sortorder]			Long Integer
);

CREATE TABLE [drawformatitem]
 (
	[id]			Long Integer, 
	[drawformat]			Long Integer, 
	[stagetype]			Long Integer, 
	[maxentries]			Long Integer, 
	[drawtype]			Long Integer, 
	[playoffsize]			Long Integer, 
	[consolation]			Long Integer, 
	[groupmethod]			Long Integer, 
	[maxroundrobinsize]			Long Integer, 
	[numtoplayoff]			Long Integer, 
	[numqual]			Long Integer
);

CREATE TABLE [eventscheduleblock]
 (
	[id]			Long Integer, 
	[eventschedule]			Long Integer, 
	[starttime]			DateTime, 
	[endtime]			DateTime, 
	[weighstart]			DateTime, 
	[weighfinish]			DateTime
);

CREATE TABLE [League]
 (
	[id]			Long Integer, 
	[name]			Text (80), 
	[eventtype]			Long Integer, 
	[gender]			Long Integer, 
	[min_age]			Long Integer, 
	[max_age]			Long Integer, 
	[minlevel]			Long Integer, 
	[maxlevel]			Long Integer, 
	[fee]			Currency, 
	[contesttime]			Long Integer, 
	[matchbreak]			Long Integer, 
	[color]			Long Integer, 
	[displayorder]			Long Integer, 
	[foreignid]			Text (50), 
	[drawformat]			Long Integer, 
	[grading]			Long Integer, 
	[goldenscoretime]			Long Integer
);

CREATE TABLE [MatchWarning]
 (
	[id]			Long Integer, 
	[match]			Long Integer, 
	[player]			Long Integer, 
	[drawname]			Text (80), 
	[playtime]			DateTime, 
	[status]			Long Integer, 
	[timestamp]			DateTime
);

CREATE TABLE [Player]
 (
	[id]			Long Integer, 
	[name]			Text (50), 
	[firstname]			Text (25), 
	[middlename]			Text (15), 
	[club]			Long Integer, 
	[country]			Long Integer, 
	[payed]			Currency, 
	[discount]			Currency, 
	[address]			Text (50), 
	[postalcode]			Text (10), 
	[city]			Text (50), 
	[state]			Text (50), 
	[phone]			Text (20), 
	[office]			Text (20), 
	[fax]			Text (20), 
	[mobile]			Text (20), 
	[email]			Text (80), 
	[gender]			Long Integer, 
	[dob]			DateTime, 
	[memberid]			Text (50), 
	[foreignid]			Text (50), 
	[usercode]			Text (50), 
	[level1]			Long Integer, 
	[level2]			Long Integer, 
	[ranking1]			Long Integer, 
	[ranking2]			Long Integer, 
	[rating1]			Double, 
	[rating2]			Double, 
	[entrydate]			DateTime, 
	[memo]			Memo/Hyperlink (255), 
	[onlineentryinfo]			Memo/Hyperlink (255), 
	[ranking3]			Long Integer, 
	[county]			Long Integer, 
	[entrymethod]			Long Integer, 
	[ranking12]			Long Integer, 
	[ranking14]			Long Integer, 
	[ranking16]			Long Integer, 
	[ranking18]			Long Integer, 
	[ranking35]			Long Integer, 
	[ranking40]			Long Integer, 
	[ranking45]			Long Integer, 
	[ranking50]			Long Integer, 
	[ranking55]			Long Integer, 
	[ranking60]			Long Integer, 
	[ranking65]			Long Integer, 
	[ranking70]			Long Integer, 
	[ranking75]			Long Integer, 
	[ranking80]			Long Integer, 
	[points1]			Double, 
	[points2]			Double, 
	[points3]			Double, 
	[address2]			Text (50), 
	[address3]			Text (50), 
	[level3]			Long Integer, 
	[rating3]			Double, 
	[acceptranking]			Long Integer, 
	[acceptproranking]			Long Integer, 
	[proranking1]			Long Integer, 
	[proranking2]			Long Integer, 
	[rankingcountry]			Long Integer, 
	[nationalranking]			Long Integer, 
	[onlineentrantid]			Long Integer, 
	[onlineentrantdate]			DateTime, 
	[playspreviousweek]			Boolean NOT NULL, 
	[bankaccount]			Text (30), 
	[bankaccountname]			Text (50), 
	[regonlineservice]			Boolean NOT NULL, 
	[rankingsi7day]			Long Integer, 
	[rankingdo7day]			Long Integer, 
	[prorankingsi7day]			Long Integer, 
	[prorankingdo7day]			Long Integer, 
	[acceptnationalranking]			Long Integer, 
	[agecheck]			Boolean NOT NULL, 
	[ipinobtained]			Boolean NOT NULL, 
	[concurrency]			Text (30), 
	[currentyearpf]			Currency, 
	[onsiteiosfee]			Currency, 
	[historicfines]			Currency, 
	[currentyearfines]			Currency, 
	[usd500fines]			Currency, 
	[totalamountdue]			Currency, 
	[minimumpayment]			Currency, 
	[itfflagshandled]			Long Integer, 
	[asianname]			Boolean NOT NULL, 
	[singlesround]			Long Integer, 
	[doublesround]			Long Integer, 
	[prizemoneysingles]			Currency, 
	[prizemoneydoubles]			Currency, 
	[taxpercentage]			Double, 
	[taxamount]			Currency, 
	[itfentryfee]			Currency, 
	[onsitefines]			Currency, 
	[outstandingfines]			Currency, 
	[otheritemspaid]			Currency, 
	[otheritemspaiddescription]			Text (50), 
	[otheritemscollected]			Currency, 
	[otheritemscollecteddescription]			Text (50), 
	[totalpaid]			Currency, 
	[otheritems]			Currency, 
	[otheritemsdescription]			Text (50), 
	[totalamount]			Currency, 
	[prizemoneycomment]			Text (100), 
	[validated]			Boolean NOT NULL, 
	[paidincash]			Currency, 
	[paymentissued]			Long Integer, 
	[prizemoneydeduction]			Currency, 
	[suspended]			Boolean NOT NULL, 
	[maxtournaments]			Boolean NOT NULL, 
	[sanctioned]			Boolean NOT NULL, 
	[ageflag]			Boolean NOT NULL, 
	[wcsinglesmain]			Long Integer, 
	[wcsinglesqual]			Long Integer, 
	[wcdoublesmain]			Long Integer, 
	[wcdoublesqual]			Long Integer, 
	[membershipcompleted]			Boolean NOT NULL, 
	[quadplayer]			Boolean NOT NULL, 
	[wcusageallowancesingles]			Long Integer, 
	[wcusageallowancedoubles]			Long Integer, 
	[lasttimeoncourt]			DateTime, 
	[acceptrankingjunior]			Long Integer, 
	[acceptrankingjunior16]			Long Integer, 
	[acceptrankingjunior14]			Long Integer, 
	[acceptrankingpro]			Long Integer, 
	[seedrankingjunior]			Long Integer, 
	[seedrankingjunior16]			Long Integer, 
	[seedrankingjunior14]			Long Integer, 
	[seedrankingpro]			Long Integer, 
	[itfrefund]			Currency, 
	[itfrefundpaid]			Currency, 
	[hastsaccount]			Boolean NOT NULL, 
	[nummatches]			Long Integer, 
	[numwins]			Long Integer, 
	[rankingstatus]			Text (25), 
	[dtbrlj]			Boolean NOT NULL, 
	[dtbrlu]			Boolean NOT NULL, 
	[dtbrla]			Boolean NOT NULL, 
	[dtbrls]			Boolean NOT NULL, 
	[dtbrgjuak]			Long Integer, 
	[dtbrgsen]			Long Integer, 
	[passport]			Boolean NOT NULL, 
	[checkedin]			Boolean NOT NULL, 
	[firstcheckin]			Boolean NOT NULL, 
	[backnr]			Long Integer, 
	[entryweight]			Double, 
	[weight]			Double, 
	[weight2]			Double, 
	[weightchecked]			Boolean NOT NULL, 
	[validationflags]			Long Integer, 
	[paraclass]			Long Integer, 
	[paraclassstatus]			Long Integer, 
	[ParaClassificationCode]			Text (50), 
	[itfsinglesstatus]			Long Integer
);

CREATE TABLE [PlayerRatingEntry]
 (
	[id]			Long Integer, 
	[ratingtype]			Long Integer, 
	[playerid]			Long Integer, 
	[rating1]			Double, 
	[rating2]			Double, 
	[rating3]			Double
);

CREATE TABLE [ScoringFormat]
 (
	[id]			Long Integer, 
	[name]			Text (100), 
	[isdefault]			Boolean NOT NULL, 
	[numsets]			Long Integer, 
	[settype]			Long Integer, 
	[lastsettype]			Long Integer, 
	[score]			Long Integer
);

CREATE TABLE [stageentry]
 (
	[ID]			Long Integer, 
	[entry]			Long Integer, 
	[stage]			Long Integer, 
	[seed1]			Long Integer, 
	[seed2]			Long Integer, 
	[status]			Long Integer, 
	[chip]			Long Integer, 
	[seedtb]			Long Integer
);


