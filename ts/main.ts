//
// Description:         Main JavaScript file for SDDS Tools Global
// Author:              Cameron Gray
// Last Modified Date:  2024-01-20
// Last Modified By:    Cameron Gray
// Version:             1.5
// 
// ##########################################################################
// NOTE: 
// This is generated from the TypeScript compiler (tsc) and is not formatted 
// as originally coded in the .ts format
// ##########################################################################
//
// Data Version: update this when the data file changes to ensure the data 
// is reloaded on clients side
const dv = "20241217.1";
const lastUpdated = "2024-12-17";
var activeTerm = "2251Global";

// Data File: CSV file containing the data to be used by the app (https://seneca-tools.github.io/ToolsGlobal/assets)
const dataFileRoot = "https://seneca-tools.github.io/ToolsGlobal/assets/";

// The fields we care about in the CSV file (assuming these field names will not change)
const CSVFieldName : string[] = [ 'Subject', 'Catalog', 'Section', 'Class Stat', 'Cap Enrl', 'Tot Enrl', 'Facil ID', 'Meeting Start',
    'Meeting End', 'Mon', 'Tues', 'Wed', 'Thurs', 'Fri', 'Mode', 'Campus', 'Instructor Name'];

// Index of CSVFieldName in CSV file (used to map the CSV fields to the object fields)
// This array is populated when the data file is loaded
// Fields can be in any order in the source CSV file
var CSVFieldIndex : number[] = [];

// Object fields mapped to the CSV fields
// These fields are used to create the object array (clientData)
const OBJFieldName : string[] = ['subject', 'catalog', 'section', 'classStat', 'capEnrol',  'totEnrol','room', 'timeStart',
    'timeEnd', 'isMon', 'isTue', 'isWed', 'isThur', 'isFri', 'mode', 'campus', 'last', 'first'];

// Overview of how the data is mapped
/*
    STATIC                  STATIC                  DYNAMIC (MAPPED)
    Object Schema           CSV Column Header       CSV Column Index
    ----------------------  --------------------    ----------------
    clientData[].courseID;    Course ID               ?
    clientData[].courseAdmin; Course Administrator    ?
    clientData[].subject;     Subject                 ?
    clientData[].catalog;     Catalog                 ?
    clientData[].section;     Section                 ?
    clientData[].classNum;    Class Nbr               ?
    clientData[].classStat;   Class Stat              ?
    clientData[].capEnrol;    Cap Enrl                ?
    clientData[].totEnrol;    Tot Enrl                ?
    clientData[].roomCap;     Room Capacity           ?
    clientData[].room;        Facil ID                ?
    clientData[].timeStart;   Meeting Start           ?
    clientData[].timeEnd;     Meeting End             ?
    clientData[].isMon;       Mon                     ?
    clientData[].isTue;       Tues                    ?
    clientData[].isWed;       Wed                     ?
    clientData[].isThu;       Thurs                   ?
    clientData[].isFri;       Fri                     ?
    clientData[].last;        Instructor Name         ?
    clientData[].first;       Instructor Name         ?
    clientData[].mode;        Mode                    ?
    clientData[].campus;      Campus                  ?
*/

// Type used for the clientData array
type ClientData = 
{
    // courseID: string;
    // courseAdmin: string;
    subject: string;
    catalog: string;
    section: string;
    // classNum: string;
    classStat: string;
    capEnrol: string;
    totEnrol: string;
    //roomCap: string;
    room: string;
    timeStart: string;
    timeEnd: string;
    isMon: string;
    isTue: string;
    isWed: string;
    isThu: string;
    isFri: string;
    mode: string;
    campus: string;
    last: string;
    first: string;
};

// Helper class for legend section results
type LegendEntry =
{
    id: string;
    backgroundColor: string;
    color: string;
    innerHTML: string;
    contactHours: number;
};

// Object Data: All records are stored in this array (on average between 600-1500 records)
var clientData = [];

// Periods and Days: Used to create the weekly schedule
type WeeklyGridDimension =
{
    maxPeriods: number;
    maxDays: number;
};

// Weekly Schedule: 16 periods per day (08:00-), 5 days per week
const weeklyGrid : WeeklyGridDimension = { maxPeriods: 16, maxDays: 5 };

// Page load event: Ensure data file is loaded into clientData array
// If the data version has changed since last loaded, reload it
window.onload = function ()
{
    if( document.getElementById("termData-select") as HTMLSelectElement != null )
    {
        (document.getElementById("termData-select") as HTMLSelectElement).value = activeTerm;

        const tryDataVersion = window.localStorage.getItem("dataVersionGlobal");
        const tryCSVFieldIndex = window.localStorage.getItem("csvFieldIndex");
        const tryClientData = window.localStorage.getItem(activeTerm);
        document.getElementById("lastUpdated").innerHTML = "Data Updated: " + lastUpdated;

        if ((tryClientData == null || tryCSVFieldIndex == null) ||
            (tryCSVFieldIndex == '[]' || tryClientData.length === 0) || tryDataVersion != dv) 
        {
            loadData();
        }
        else 
        {
            CSVFieldIndex = JSON.parse(tryCSVFieldIndex);
            clientData = JSON.parse(tryClientData);
        }
    }
};

// Data Source Change Event: Reload the data file when the data source has changed
function dataSourceChange()
{
    activeTerm = (document.getElementById("termData-select") as HTMLSelectElement).value;

    const tryClientData = window.localStorage.getItem(activeTerm);

    if ((tryClientData == null || tryClientData.length === 0)) 
    {
        loadData();
    }
    else 
    {
        clientData = JSON.parse(tryClientData);
    }
}
// Load the data file into the clientData array
// This function is called when the data file is not found in local storage or
// the data version has changed
function loadData() 
{
    clientData.length = 0; // Clear the array
    CSVFieldIndex.length = 0; // Clear the array
    var dataFile = dataFileRoot + activeTerm + ".csv"; 

    console.debug("Loading data file: " + dataFile);

    // Read the file
    fetch(dataFile).then(function (response) 
    {
        return response.text();
    }).then(function (text) 
    {
        var rawData = [];
        var i = 0, max;
        // Split the file into an array of lines
        rawData = text.split('\n');
        max = rawData.length;
        // Contains data if it contains header row and 1 or more data rows
        if (max > 2) 
        {
            // Get the header row
            // Remove all csv quotes and commas and replace with tabs
            // so we can reliably determine each field by splitting on tab char
            var csvHDR = rawData[0];
            csvHDR = csvHDR.replace(/","/g, '\t');
            csvHDR = csvHDR.replace(/"/g, '');

            // Split the header row into an array of fields
            var rawFields = [];
            rawFields = csvHDR.split('\t');

            // Map the fields to respective index for only the fields we care about
            for (let i = 0; i < rawFields.length; i++) 
            {
                var idx = CSVFieldName.indexOf(rawFields[i]);
                if (idx >= 0) 
                {
                    // Map the field name to the index
                    CSVFieldIndex[idx] = i;

                    // The first/last names are in one field in the source csv
                    // Share the 'Instructor Name' field with object's last and first fields
                    if (OBJFieldName[idx] == 'last') 
                    {
                        CSVFieldIndex[idx + 1] = i;
                    }
                }
            }

            // Start on the first actual data row and process each row
            for (i = 1; i < max - 1; i++) 
            {
                var row = rawData[i];

                // Remove all csv quotes and commas and replace with tabs
                // so we can reliably determine each field by splitting on tab char
                row = row.replace(/","/g, '\t');
                row = row.replace(/"/g, '');

                // Split the row into an array of fields
                var rawFields = [];
                rawFields = row.split('\t');

                // A single record
                const clientDataRec : ClientData = 
                {
                    // courseID: rawFields[CSVFieldIndex[OBJFieldName.indexOf('courseID')]],
                    // courseAdmin: rawFields[CSVFieldIndex[OBJFieldName.indexOf('courseAdmin')]],
                    subject: rawFields[CSVFieldIndex[OBJFieldName.indexOf('subject')]],
                    catalog: rawFields[CSVFieldIndex[OBJFieldName.indexOf('catalog')]].replace(/ /g, ''),
                    section: rawFields[CSVFieldIndex[OBJFieldName.indexOf('section')]],
                    // classNum: rawFields[CSVFieldIndex[OBJFieldName.indexOf('classNum')]],
                    classStat: rawFields[CSVFieldIndex[OBJFieldName.indexOf('classStat')]],
                    capEnrol: rawFields[CSVFieldIndex[OBJFieldName.indexOf('capEnrol')]],
                    totEnrol: rawFields[CSVFieldIndex[OBJFieldName.indexOf('totEnrol')]],
                    // roomCap: rawFields[CSVFieldIndex[OBJFieldName.indexOf('roomCap')]],
                    room: rawFields[CSVFieldIndex[OBJFieldName.indexOf('room')]],
                    timeStart: rawFields[CSVFieldIndex[OBJFieldName.indexOf('timeStart')]],
                    timeEnd: rawFields[CSVFieldIndex[OBJFieldName.indexOf('timeEnd')]],
                    isMon: rawFields[CSVFieldIndex[OBJFieldName.indexOf('isMon')]],
                    isTue: rawFields[CSVFieldIndex[OBJFieldName.indexOf('isTue')]],
                    isWed: rawFields[CSVFieldIndex[OBJFieldName.indexOf('isWed')]],
                    isThu: rawFields[CSVFieldIndex[OBJFieldName.indexOf('isThur')]],
                    isFri: rawFields[CSVFieldIndex[OBJFieldName.indexOf('isFri')]],
                    mode: rawFields[CSVFieldIndex[OBJFieldName.indexOf('mode')]],
                    campus: rawFields[CSVFieldIndex[OBJFieldName.indexOf('campus')]],
                    last: "",
                    first: ""
                };

                // Instructor Name: split into last and first
                var nameParts = [];
                nameParts = rawFields[CSVFieldIndex[OBJFieldName.indexOf('last')]].split(',');
                clientDataRec.last = nameParts[0];

                if (nameParts.length > 1) 
                {
                    clientDataRec.first = nameParts[1];
                }
                else 
                {
                    clientDataRec.first = '';
                }

                // Store it to global data object array
                clientData.push(clientDataRec);
            }
        }

        // Store the data in local storage for persistence
        // and to avoid having to re-read the file each time
        // these values are reassessed each time a page is loaded
        // to ensure the latest data is being used at all times
        window.localStorage.setItem("csvFieldIndex", JSON.stringify(CSVFieldIndex));
        window.localStorage.setItem(activeTerm, JSON.stringify(clientData));
        window.localStorage.setItem("dataVersionGlobal", dv);

    }).catch(function (err) 
    {
        dbgInfo('Error: Failed to read file: ' + err);
    });
}

// Quick and dirty function to get a unique colour based on the value passed in
// This is used to colour the object in the schedule and legend
function getColour(val) 
{
    // Default colour scheme (object representing text colour and background colour)
    var retValue = { background: "grey", color: "black" };
    var alpha = 1.0;

    // If the value is greater than 19, then we need to use a different alpha value
    // to make the colour lighter
    // this logic only supports up to 38 unique colours
    if (val > 19) 
    {
        val %= 19;
        alpha = 0.6;
    }

    // assign the colour based on the value passed in
    switch (val) 
    {
        case 0: // red
            retValue.background = "rgba(230, 25, 75, " + alpha + ")";
            retValue.color = "white";
            break;
        case 1: // green
            retValue.background = "rgba(60, 180, 75,  " + alpha + ")";
            retValue.color = "white";
            break;
        case 2: // yellow
            retValue.background = "rgba(255, 225, 25, " + alpha + ")";
            break;
        case 3: // blue
            retValue.background = "rgba(0, 130, 200, " + alpha + ")";
            retValue.color = "white";
            break;
        case 4: // orange
            retValue.background = "rgba(245, 130, 48, " + alpha + ")";
            retValue.color = "white";
            break;
        case 5: // purple
            retValue.background = "rgba(145, 30, 180, " + alpha + ")";
            retValue.color = "white";
            break;
        case 6: // cyan
            retValue.background = "rgba(70, 240, 240, " + alpha + ")";
            break;
        case 7: // magenta
            retValue.background = "rgba(240, 50, 230, " + alpha + ")";
            retValue.color = "white";
            break;
        case 8: // lime
            retValue.background = "rgba(210, 245, 60, " + alpha + ")";
            break;
        case 9: // pink
            retValue.background = "rgba(250, 190, 212, " + alpha + ")";
            break;
        case 10: // teal
            retValue.background = "rgba(0, 128, 128, " + alpha + ")";
            retValue.color = "white";
            break;
        case 11: // lavender
            retValue.background = "rgba(220, 190, 255, " + alpha + ")";
            break;
        case 12: // brown
            retValue.background = "rgba(170, 110, 40, " + alpha + ")";
            retValue.color = "white";
            break;
        case 13: // beige
            retValue.background = "rgba(255, 250, 200, " + alpha + ")";
            break;
        case 14: // maroon
            retValue.background = "rgba(128, 0, 0, " + alpha + ")";
            retValue.color = "white";
            break;
        case 15: // mint
            retValue.background = "rgba(170, 255, 195, " + alpha + ")";
            break;
        case 16: // olive
            retValue.background = "rgba(128, 128, 0, " + alpha + ")";
            retValue.color = "white";
            break;
        case 17: // coral/apricot
            retValue.background = "rgba(255, 215, 180, " + alpha + ")";
            break;
        case 18: // navy
            retValue.background = "rgba(0, 0, 128, " + alpha + ")";
            retValue.color = "white";
            break;
        case 19: // grey
            retValue.background = "rgba(128, 128, 128, " + alpha + ")";
            retValue.color = "white";
            break;
    }

    return retValue;
}

// FACULTY workhorse function to create filter and display the schedule results
function processFaculty() 
{
    // reset the schedule (blank it out)
    reset();

    // create the schedule table
    makeWeeklySchedule();

    // clear the contents of the debug element
    document.getElementById('debugInfo').innerHTML = '';

    // get the search string from the input field
    var search = (document.getElementById('queryValue') as HTMLTextAreaElement).value;

    // only process if there is a search string specified
    // and is not exceedingly long
    if (search.length > 0 && search.length < 1500) 
    {
        // split the search string into individual faculty names
        var facultyList = search.split('|');

        // faculty counter and indexer
        // used to assign a colour to each faculty found
        var facultyCount = -1;

        // colour to use for each course section (default to white/black)
        var colour = { color: 'black', background: 'white' };

        // legend listing of faculty
        var legendList : LegendEntry[] = [];

        // loop through the data: 
        for (let i = 0; i < clientData.length; i++) 
        {
            // current data record
            var rec = clientData[i];

            // loop through the faculty list:
            for (let j = 0; j < facultyList.length; j++) 
            {
                // extract the last and first name parts
                const qryFaculty = facultyList[j];
                var [searchLast, searchFirst] = qryFaculty.split(/\s*,\s*/);

                if (searchLast === undefined && searchFirst === undefined) 
                {
                    // hate doing this... but it makes sense here...
                    // Just ignore the entry if it is blank
                    continue;
                }
                // Workaround for the case where there is no last name supplied
                // Permit lookups by first name only
                if (searchLast === undefined) 
                {
                    searchLast = "";
                }
                // Workaround for the case where there is no first name supplied
                // Permit lookups by last name only
                if (searchFirst === undefined) 
                {
                    searchFirst = "";
                }

                // Remove any leading or trailing spaces
                searchLast = searchLast.trim();
                searchFirst = searchFirst.trim();

                // Guard against blank entries
                if (searchLast.length == 0 && searchFirst.length == 0) 
                {
                    // again.. hate doing this... but it makes sense here...
                    // Just ignore the entry if both name parts are blank
                    continue;
                }

                // Do we have a match?
                // This is a case insensitive search and will match on any part of the name
                if (rec.last.toLowerCase().includes(searchLast.toLowerCase()) &&
                    rec.first.toLowerCase().includes(searchFirst.toLowerCase())) 
                {
                    // Do we already have this faculty in the legend listing?
                    var idx = findLegendEntry(legendList, rec.last + rec.first);

                    if( idx < 0)
                    {
                        // increment the faculty counter
                        facultyCount++;

                        // assign a colour to the faculty
                        colour = getColour(facultyCount);

                        var entry : LegendEntry = {
                            // Make a unique id for the legend item so we can find it later and prevent duplicates
                            id: rec.last + rec.first,
                            backgroundColor: colour.background,
                            color: colour.color,
                            innerHTML: rec.last + ', ' + rec.first,
                            contactHours: 0
                        };

                        // add to the array of legend entries
                        legendList.push(entry);

                        // update idx reference to new position
                        idx = facultyCount;
                    }
                    else
                    {
                        colour.background = legendList[idx].backgroundColor;
                        colour.color = legendList[idx].color;
                    }

                    // Need to span multiple rows if the class spans more than one period (55 minutes)
                    const periodDuration = getStartPeriodDuration(rec.timeStart, rec.timeEnd);
                    var day = getDayOfWeek(rec);
                    day.forEach(wkDay => 
                    {
                        // accumulate the contact hours (1 period = 1 contact hour)
                        legendList[idx].contactHours += periodDuration.duration;

                        // Schedule cell data object container
                        const section = document.createElement("div");
                        section.classList.add('FacultyLabel');
                        section.style.backgroundColor = legendList[idx].backgroundColor;
                        section.style.borderColor = legendList[idx].backgroundColor;
                        section.style.color = legendList[idx].color;
                        section.innerHTML = rec.subject + rec.catalog + ' (' + rec.section + ')<br />' + rec.first[0] + '. ' + rec.last + ' (enrl:' + rec.totEnrol + ')<br />' + rec.room;
                        section.title = rec.subject + rec.catalog + ':' + rec.section + '\n' + rec.first[0] + '.' + rec.last + '\n' + rec.room + '\n' + rec.timeStart + '-' + rec.timeEnd + '\n' + rec.campus;
                        
                        // Add the data cell to the schedule table
                        addElement(periodDuration, wkDay, section, colour);
                    });
                }
            }
        }

        // Maintain legend section
        if( legendList.length > 0 )
        {
            // Sort the legend list
            legendList.sort((a, b) => (a.id > b.id) ? 1 : -1);

            for(let i = 0; i < legendList.length; i++)
            {
                // legend item container
                const box = document.createElement("div");

                // Make a unique id for the legend item so we can find it later and prevent duplicates
                box.id = legendList[i].id;
                box.className = 'legendItem';
                box.style.backgroundColor = legendList[i].backgroundColor;
                box.style.color = legendList[i].color;

                // legend item text
                const facName = document.createElement("span");
                facName.style.backgroundColor =legendList[i].backgroundColor;
                facName.style.color = legendList[i].color;
                facName.innerHTML = legendList[i].innerHTML + ' (' + legendList[i].contactHours + ')';
                box.appendChild(facName);

                // add to the main legend element
                document.getElementById('legend').appendChild(box);
            }
        }
    }
}

// COURSE workhorse function to create filter and display the schedule results
function processCourse() 
{
    // reset the schedule (blank it out)
    reset();

    // create the schedule table
    makeWeeklySchedule();

    // clear the contents of the debug element
    document.getElementById('debugInfo').innerHTML = '';

    // get the search string from the input field
    var search = (document.getElementById('queryValue') as HTMLTextAreaElement).value;

    // only process if there is a search string specified
    // and is within reasonable length (1 course : 6 chars)
    if (search.length > 0 && search.length < 10)
    {
        // extract the course and course number parts
        var course = search.substring(0, 3);
        var courseNum = search.substring(3);

        // Course section counter and indexer
        // used to assign a colour to each course section
        var courseSectionCount = -1;

        // legend listing of course sections
        var courseSectionList = [];

        // colour to use for each course section (default to white/black)
        var colour = { color: 'black', background: 'white' };

        // legend listing of course/faculty
        var legendList : LegendEntry[] = [];

        // Loop through the data
        for (let i = 0; i < clientData.length; i++) 
        {
            // current data record
            var rec = clientData[i];

            // Do we have a match?
            // This is a case insensitive search and will match on any part of the subject and catalog parts
            // Only display courses designated with fewer than 50 seats and is active
            // (This is to filter out the "dummy" courses that are used for scheduling purposes etc...)
            if (!(rec.room == '') && !(rec.last == '') &&
                (rec.classStat == 'A' && rec.capEnrol < 50) &&
                rec.subject.toLowerCase().includes(course.toLowerCase()) &&
                rec.catalog.toLowerCase().includes(courseNum.toLowerCase())) 
            {
                // Do we already have this faculty in the legend listing?
                var idx = findLegendEntry(legendList, rec.subject + rec.catalog + rec.section.substring(0, 3) + rec.first[0] + '. ' + rec.last);

                if( idx < 0)
                {
                    // increment the course section counter (used to assign a colour to each course section)
                    courseSectionCount++;
                    colour = getColour(courseSectionCount);

                    var entry : LegendEntry = {
                        // Make a unique id for the legend item so we can find it later and prevent duplicates
                        id:rec.subject + rec.catalog + rec.section.substring(0, 3) + rec.first[0] + '. ' + rec.last,
                        backgroundColor: colour.background,
                        color: colour.color,
                        innerHTML: rec.subject.substring(0, 3) + rec.catalog + ' - ' + rec.section.substring(0, 3) + ' (enrl:' + rec.totEnrol +
                                    ')<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(' + rec.last + ', ' + rec.first + ')',
                        contactHours: 0
                    };

                    // add to the array of legend entries
                    legendList.push(entry);

                    // update idx reference to new position
                    idx = courseSectionCount;                    
                }
                else
                {
                    colour.background = legendList[idx].backgroundColor;
                    colour.color = legendList[idx].color;
                }

                // Need to span multiple rows if the class spans more than one period (55 minutes)
                const periodDuration = getStartPeriodDuration(rec.timeStart, rec.timeEnd);
                var day = getDayOfWeek(rec);
                day.forEach(wkDay => 
                {
                    // Schedule cell data object container
                    const section = document.createElement("div");
                    section.classList.add('FacultyLabel');
                    section.style.backgroundColor = legendList[idx].backgroundColor;
                    section.style.borderColor = legendList[idx].backgroundColor;
                    section.style.color = legendList[idx].color;
                    section.innerHTML = rec.subject + rec.catalog + ' (' + rec.section + ')<br />' + rec.first[0] + '. ' +
                                         rec.last + '<br />' + rec.room;
                    section.title = rec.subject + rec.catalog + ':' + rec.section + '\n' + rec.first[0] + '.' + rec.last + 
                                        '\n' + rec.room + '\n' + rec.timeStart + '-' + rec.timeEnd + '\n' + rec.campus;

                    // Add the data cell to the schedule table
                    addElement(periodDuration, wkDay, section, colour);
                });
            }
        }

        // Maintain legend section
        if( legendList.length > 0 )
        {
            // Sort the legend list
            legendList.sort((a, b) => (a.id > b.id) ? 1 : -1);

            for(let i = 0; i < legendList.length; i++)
            {
                // legend item container
                const box = document.createElement("div");

                // Make a unique id for the legend item so we can find it later and prevent duplicates
                box.id = legendList[i].id;
                box.className = 'legendItem';
                box.style.backgroundColor = legendList[i].backgroundColor;
                box.style.color = legendList[i].color;

                // legend item text
                const facName = document.createElement("span");
                facName.style.backgroundColor =legendList[i].backgroundColor;
                facName.style.color = legendList[i].color;
                facName.innerHTML = legendList[i].innerHTML;
                box.appendChild(facName);

                // add to the main legend element
                document.getElementById('legend').appendChild(box);
            }
        }        
    }
}

function processRoom() 
{
    // reset the schedule (blank it out)
    reset();

    // create the schedule table
    makeWeeklySchedule();

    // clear the contents of the debug element
    document.getElementById('debugInfo').innerHTML = '';

    // get the search string from the input field
    var search = (document.getElementById('queryValue') as HTMLTextAreaElement).value;

    // only process if there is a search string specified
    // and is within reasonable length (1 room : 5 chars)
    if (search.length > 0 && search.length < 10) 
    {
        // work in upper case
        var room = search.toUpperCase();

        // Course section counter and indexer
        // used to assign a colour to each course section found
        var courseSectionCount = -1;

        // legend listing of course sections
        var courseSectionList = [];

        // colour to use for each course section (default to white/black)
        var colour = { color: 'black', background: 'white' };

        // legend listing of course/faculty
        var legendList : LegendEntry[] = [];

        // Loop through the data
        for (let i = 0; i < clientData.length; i++) 
        {
            // current data record
            var rec = clientData[i];

            // Do we have a match?
            // This is a case insensitive search and will match on any part of the subject and catalog parts
            // Only display courses designated with fewer than 50 seats and is active
            // (This is to filter out the "dummy" courses that are used for scheduling purposes etc...)
            if (!(rec.room == '') && !(rec.last == '') &&
                (rec.classStat == 'A' && rec.capEnrol < 50) &&
                rec.room.toUpperCase().includes(room)) 
            {

                // Do we already have this faculty in the legend listing?
                var idx = findLegendEntry(legendList, rec.room + rec.subject + rec.catalog + rec.section.substring(0, 3) + rec.first[0] + '. ' + rec.last);

                if( idx < 0)
                {
                    // increment the course section counter (used to assign a colour to each course section)
                    courseSectionCount++;
                    colour = getColour(courseSectionCount);

                    var entry : LegendEntry = {
                        // Make a unique id for the legend item so we can find it later and prevent duplicates
                        id: rec.room + rec.subject + rec.catalog + rec.section.substring(0, 3) + rec.first[0] + '. ' + rec.last,
                        backgroundColor: colour.background,
                        color: colour.color,
                        innerHTML: rec.room + ' : ' + rec.subject.substring(0, 3) + rec.catalog + ' - ' + rec.section.substring(0, 3) + '(enrl:' + rec.totEnrol +
                                    ')<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + rec.last + ', ' + rec.first,
                        contactHours: 0
                    };

                    // add to the array of legend entries
                    legendList.push(entry);

                    // update idx reference to new position
                    idx = courseSectionCount;                    
                }
                else
                {
                    colour.background = legendList[idx].backgroundColor;
                    colour.color = legendList[idx].color;
                }

                // Need to span multiple rows if the class spans more than one period (55 minutes)
                const periodDuration = getStartPeriodDuration(rec.timeStart, rec.timeEnd);
                var day = getDayOfWeek(rec);
                day.forEach(wkDay => 
                {
                    // Schedule cell data object container
                    const section = document.createElement("div");
                    section.classList.add('FacultyLabel');
                    section.style.backgroundColor = legendList[idx].backgroundColor;
                    section.style.borderColor = legendList[idx].backgroundColor;
                    section.style.color = legendList[idx].color;
                    section.innerHTML = rec.subject + rec.catalog + ' (' + rec.section + ')<br />' + rec.first[0] + '. ' +
                                        rec.last + '<br />' + rec.room;
                    section.title = rec.subject + rec.catalog + ':' + rec.section + '\n' + rec.first[0] + '.' + 
                                    rec.last + '\n' + rec.room + '\n' + rec.timeStart + '-' + rec.timeEnd + '\n' + rec.campus;

                    // Add the data cell to the schedule table
                    addElement(periodDuration, wkDay, section, colour);
                });
            }
        }

        // Maintain legend section
        if( legendList.length > 0 )
        {
            // Sort the legend list
            legendList.sort((a, b) => (a.id > b.id) ? 1 : -1);

            for(let i = 0; i < legendList.length; i++)
            {
                // legend item container
                const box = document.createElement("div");

                // Make a unique id for the legend item so we can find it later and prevent duplicates
                box.id = legendList[i].id;
                box.className = 'legendItem';
                box.style.backgroundColor = legendList[i].backgroundColor;
                box.style.color = legendList[i].color;

                // legend item text
                const facName = document.createElement("span");
                facName.style.backgroundColor =legendList[i].backgroundColor;
                facName.style.color = legendList[i].color;
                facName.innerHTML = legendList[i].innerHTML;
                box.appendChild(facName);

                // add to the main legend element
                document.getElementById('legend').appendChild(box);
            }
        }        
    }
}

// Set text to the reserved element for debugging information
function dbgInfo(msg, addNewline = true) 
{
    document.getElementById('debugInfo').innerHTML += msg;

    if (addNewline) 
    {
        document.getElementById('debugInfo').innerHTML += '<br>\n';
    }
}

// Get the day of week numeric value from the data record (1 = Monday, 5 = Friday)
function getDayOfWeek(obj) 
{
    var day = [];

    if (obj.isMon == 'Y') { day.push(1); }
    if (obj.isTue == 'Y') { day.push(2); }
    if (obj.isWed == 'Y') { day.push(3); }
    if (obj.isThu == 'Y') { day.push(4); }
    if (obj.isFri == 'Y') { day.push(5); }

    return day;
}

// Get the numeric start row and end row for the given string time value range
// used in spanning the data cell over multiple rows
function getStartPeriodDuration(timeStart, timeEnd) 
{
    var startPeriodRow = 0;
    var durationRows = 0;
    var startPeriod = timeStart.split(':');
    var startHour = parseInt(startPeriod[0]);
    var startMin = parseInt(startPeriod[1]);
    var endPeriod = timeEnd.split(':');
    var endHour = parseInt(endPeriod[0]);
    var endMin = parseInt(endPeriod[1]);

    if (startHour >= 8) 
    {
        // NOTE: javascript does not support this arithmetic expression! 
        //       had to do this in multiple steps!
        //startPeriodRow = (startHour * 60 + startMin) - 480.0 / 55.0 + 1;
        startPeriodRow = startHour * 60 + startMin;
        startPeriodRow = startPeriodRow - 480.0;
        startPeriodRow = startPeriodRow / 55.0;
        startPeriodRow = startPeriodRow + 1;
    }

    // NOTE: javascript does not support this arithmetic expression! 
    //       had to do this in multiple steps!
    //durationRows = (endHour * 60 + endMin) - (startHour * 60 + startMin);
    durationRows = endHour * 60;
    durationRows = durationRows + endMin;
    durationRows = durationRows - (startHour * 60);
    durationRows = durationRows - startMin;
    durationRows = durationRows + 5;
    durationRows = durationRows / 55.0;
    const ret = { period: startPeriodRow, duration: durationRows };

    return ret;
}

// Get the numeric start row for the given string time value
function getStartPeriodRow(timeStart) 
{
    let startPeriodRow = 0;
    var startPeriod = timeStart.split(':');
    var startHour = parseInt(startPeriod[0]);
    var startMin = parseInt(startPeriod[1]);

    if (startHour >= 8) 
    {
        // NOTE: javascript does not support this arithmetic expression! 
        //       had to do this in multiple steps!
        //startPeriodRow = (((startHour * 60 + startMin) - 480.0) / 55.0) + 1;
        startPeriodRow = startHour * 60 + startMin;
        startPeriodRow = startPeriodRow - 480.0;
        startPeriodRow = startPeriodRow / 55.0;
        startPeriodRow += 1;
    }

    return startPeriodRow;
}

// Get the numeric number of rows to span for the given string time value range
function getDurationRows(timeStart, timeEnd) 
{
    var durationRows = 0;
    var startPeriod = timeStart.split(':');
    var startHour = parseInt(startPeriod[0]);
    var startMin = parseInt(startPeriod[1]);
    var endPeriod = timeEnd.split(':');
    var endHour = parseInt(endPeriod[0]);
    var endMin = parseInt(endPeriod[1]);

    if (startHour >= 8) 
    {
        // NOTE: javascript does not support this arithmetic expression! 
        //       had to do this in multiple steps!
        //durationRows = (endHour * 60 + endMin) - (startHour * 60 + startMin) + 5;
        //durationRows = durationRows / 55.0;
        durationRows = endHour * 60;
        durationRows = durationRows + endMin;
        durationRows = durationRows - (startHour * 60);
        durationRows = durationRows + startMin;
        durationRows = durationRows / 55.0;
    }

    return durationRows;
}

// Create the main weekly schedule results table
// contains header row and column titles, and data cell containers for storing the results
function makeWeeklySchedule() 
{
    // Main container element for the weekly schedule table
    const container = document.getElementById("weekView");

    // divide the container into a grid of rows and columns
    container.style.setProperty('--grid-rows', weeklyGrid.maxPeriods.toString());
    container.style.setProperty('--grid-cols', weeklyGrid.maxDays.toString());
    
    let x :number, y : number;

    // Create the header row for the weekly day titles
    for (x = 0; x < weeklyGrid.maxDays + 1; x++) 
    {
        var weekday = document.createElement("h4");
        weekday.classList.add("grid-item", "DayAxis");

        switch (x) {
            case 0:
                weekday.innerHTML = "Time/Day";
                break;
            case 1:
                weekday.innerHTML = "MONDAY";
                break;
            case 2:
                weekday.innerHTML = "TUESDAY";
                break;
            case 3:
                weekday.innerHTML = "WEDNESDAY";
                break;
            case 4:
                weekday.innerHTML = "THURSDAY";
                break;
            case 5:
                weekday.innerHTML = "FRIDAY";
                break;
        }
        container.appendChild(weekday);
    }

    // Create the header column for the daily time titles
    for (x = 1; x < weeklyGrid.maxPeriods + 1; x++) 
    {
        for (y = 0; y < weeklyGrid.maxDays + 1; y++) 
        {
            if (y == 0) 
            {
                let cell = document.createElement("h4");
                cell.classList.add("grid-item", "TimeAxis");

                if (x == 1) { cell.innerHTML = "8:00"; }
                else if (x == 2) { cell.innerHTML = "8:55"; }
                else if (x == 3) { cell.innerHTML = "9:50"; }
                else if (x == 4) { cell.innerHTML = "10:45"; }
                else if (x == 5) { cell.innerHTML = "11:40"; }
                else if (x == 6) { cell.innerHTML = "12:35"; }
                else if (x == 7) { cell.innerHTML = "1:30"; }
                else if (x == 8) { cell.innerHTML = "2:25"; }
                else if (x == 9) { cell.innerHTML = "3:20"; }
                else if (x == 10) { cell.innerHTML = "4:15"; }
                else if (x == 11) { cell.innerHTML = "5:10"; }
                else if (x == 12) { cell.innerText = "6:05"; }
                else if (x == 13) { cell.innerHTML = "7:00"; }
                else if (x == 14) { cell.innerHTML = "7:55"; }
                else if (x == 15) { cell.innerHTML = "8:50"; }
                else if (x == 16) { cell.innerHTML = "9:45"; }
                
                container.appendChild(cell); //.className = "grid-item";
            }

            // Create the data cell containers for the data to be stored in
            else 
            {
                let cell = document.createElement("div");

                // unique id for each cell consisting of the period and day parts
                // ex: p1d1
                cell.id = "p" + (x) + "d" + (y);
                cell.classList.add("grid-item", "DataCell");

                container.appendChild(cell);
            }
        }
    }
}

// Add the given element to the weekly schedule table coordinates
// period: period.period is the starting row
//         period.duration is the number of rows to span
// day   : is the column of the cell container
// obj   : is the parent element to use as a basis for the new data element
// col   : is the color object to use for the element
function addElement(period, day, obj, col) 
{
    const contain = document.getElementById("p" + parseInt(period.period) + "d" + parseInt(day));

    obj.classList.add("popup");
    obj.onclick = function () 
    {
        openPopup(obj);
    };

    var popMain = document.createElement("span");
    popMain.classList.add("popuptext");
    popMain.innerHTML = obj.title.replace(/\n/g, "<br>");
    obj.appendChild(popMain);
    contain.appendChild(obj);

    for (let i = 1; i < period.duration; i++) 
    {
        if (period.period + i < weeklyGrid.maxPeriods + 1)  //14) 
        {
            var sub = document.createElement("div");
            sub.innerHTML = obj.innerHTML;
            sub.title = obj.title;
            sub.style.backgroundColor = col.background;
            sub.style.borderColor = col.background;
            sub.style.color = col.color;
            sub.classList.add("FacultyLabel");
            sub.classList.add("popup");
            sub.onclick = function () 
            {
                openPopup(sub);
            };

            var pop = document.createElement("span");
            pop.classList.add("popuptext");
            pop.innerHTML = obj.title.replace(/\n/g, "<br>");

            sub.appendChild(pop);
            const container = document.getElementById("p" + parseInt(period.period + i) + "d" + parseInt(day));
            container.appendChild(sub);
        }
    }
}

// Display the objects child as a popup element (when data cell is clicked)
function openPopup(obj) 
{
    var popup = obj.getElementsByClassName("popuptext")[0];
    popup.classList.toggle("show");
}

// Clear the weekly view schedule and legend components
function reset() 
{
    const container = document.getElementById("weekView");
    while (container.firstChild) 
    {
        container.removeChild(container.firstChild);
    }
    
    const legend = document.getElementById("legend");
    while (legend.firstChild) 
    {
        legend.removeChild(legend.firstChild);
    }
}

// Search a LegendEntry array and return zero-based index
function findLegendEntry(legend:LegendEntry[], val:string)
{
    for (let i = 0; i < legend.length; i++) 
    {
        if (legend[i].id == val) 
        {
            return i;
        }
    }
    return -1;
}