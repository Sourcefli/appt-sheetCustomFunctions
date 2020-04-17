
/**
* Summary - Get total pre-sets by entering the agent's name, the starting date range, and the ending date range
* 
* 
* @example
* returns Rich Schlemmers' Total Preset Sales From April 1st to April 30th
* GETAGENTPRESETSALES("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function GETAGENTPRESETSALES(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsOnlyNicoleTeamPresets = ArrayLib.filterByText(filteredApptsByAgentName, 3, "Nicole OX")
  const filteredApptsOnlyLoisaTeamPresets = ArrayLib.filterByText(filteredApptsByAgentName, 3, ["Loisa LP", "Gee LP", "Reyna LP", "Jenny LP","Sweet LP", "Mitch LP"])
  const filteredApptsBySoldNicolesTeam = ArrayLib.filterByText(filteredApptsOnlyNicoleTeamPresets, 16, 'Sold')
  const filteredApptsBySoldLoisasTeam = ArrayLib.filterByText(filteredApptsOnlyLoisaTeamPresets, 16, 'Sold')
  const totalApptCount = 0 + filteredApptsBySoldNicolesTeam.length + filteredApptsBySoldLoisasTeam.length
  return totalApptCount
}



/**
* Summary - Get total pre-sets assigned to an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* returns Rich Schlemmers' Total Assigned Presets From April 1st to April 30th
* GETAGENTSTOTALPRESETS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function GETAGENTSTOTALPRESETS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByAgentName, 4, ["PCFE1", 'T65-1-NV-CA', 'T65-2-NV-AZ', 'AS1', 'AS2', 'AS3'])
//  const filteredApptsOnlyLoisaTeamPresets = ArrayLib.filterByText(filteredApptsByAgentName, 4, ["Loisa LP", "Gee LP", "Reyna LP", "Jenny LP","Sweet LP", "Mitch LP"])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}

/* ===================== TESTING ========================= */

function run() {
//  var res = GETAGENTPRESETSALES("RICHS", "04/12/2020","04/20/2020")
  var res = GETAGENTSTOTALPRESETS("RICHS", "04/12/2020","04/20/2020")
  Logger.log(res)
}



