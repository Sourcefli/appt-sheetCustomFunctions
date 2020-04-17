
/**
* Summary - Get total pre-sets by entering the agent's name, the starting date range, and the ending date range
* 
* 
* @example
* returns Rich Schlemmers' Total Preset Sales From April 1st to April 30th
* AGENT_TOTAL_PRESETS_SOLD("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_PRESETS_SOLD(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByAgentName, 4, ["PCFE1", 'T65-1-NV-CA', 'T65-2-NV-AZ', 'AS1', 'AS2', 'AS3'])
  const filteredApptsBySold = ArrayLib.filterByText(filteredApptsByPresetCampaigns, 16, 'Sold')
  const totalApptCount = 0 + filteredApptsBySold.length
  return totalApptCount
}



/**
* Summary - Get total pre-sets assigned to an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* returns Rich Schlemmers' Total Assigned Presets From April 1st to April 30th
* AGENT_TOTAL_PRESETS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_PRESETS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByAgentName, 4, ["PCFE1", 'T65-1-NV-CA', 'T65-2-NV-AZ', 'AS1', 'AS2', 'AS3'])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}



/**
* Summary - Get total medicare related pre-sets assigned to an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* Returns Total Medicare Related Presets For Rich Schlemmer From April 1st to April 30th
* AGENT_TOTAL_PRESETS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_MEDICARE_PRESETS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByAgentName, 4, ['T65-1-NV-CA', 'T65-2-NV-AZ', 'AS1', 'AS2', 'AS3'])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}



/**
* Summary - Get total final expense related pre-sets assigned to an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* Returns Total Final Expense Related Presets For Rich Schlemmer From April 1st to April 30th
* AGENT_TOTAL_FINALEXPENSE_PRESETS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_FINALEXPENSE_PRESETS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByAgentName, 4, ['PCFE1'])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}




/**
* Summary - Get total Medicare campaign related interviews (i.e. 'sold','no sale', 'follow up') for an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* Returns Total Medicare Related Preset Interviews Conducted by Rich Schlemmer From April 1st to April 30th
* AGENT_TOTAL_MEDICARE_INTERVIEWS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_MEDICARE_INTERVIEWS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByDisposition = ArrayLib.filterByText(filteredApptsByAgentName, 16, ['sold', 'no sale', 'follow up'])
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByDisposition, 16, ['T65-1-NV-CA', 'T65-2-NV-AZ', 'AS1', 'AS2', 'AS3'])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}



/**
* Summary - Get total Final Expense related interviews (i.e. 'sold','no sale', 'follow up') for an agent by entering their name, the starting date, and the ending date
* 
* 
* @example
* Returns Total Final Expense Related Preset Interviews Conducted by Rich Schlemmer From April 1st to April 30th
* AGENT_TOTAL_FINALEXPENSE_INTERVIEWS("RICHS", "04/01/2020", "04/30/2020")
*
* @param {agentName} - The agents name you'd like to filter by. Use their First Name, Last Initial => All Caps, no spaces (e.g. "RICHS")
* @param {startDate} - The starting date you'd like to filter by
* @param {endDate} - The ending date you'd like to filter by
*
* @return {number} - The total sales for this agent, within the given time range
* @customfunction
*/
function AGENT_TOTAL_FINALEXPENSE_INTERVIEWS(agentName, startDate, endDate) {
  const SS = SpreadsheetApp.getActiveSpreadsheet()
  const SHT = SS.getSheetByName('April2020DataImport')
  const allData = SHT.getDataRange().getValues()
  const filteredApptsByDateRange = filterApptsByDateRange(allData, startDate, endDate)
  const filteredApptsByAgentName = ArrayLib.filterByText(filteredApptsByDateRange, 14, agentName)
  const filteredApptsByDisposition = ArrayLib.filterByText(filteredApptsByAgentName, 16, ['sold', 'no sale', 'follow up'])
  const filteredApptsByPresetCampaigns = ArrayLib.filterByText(filteredApptsByDisposition, 16, ['PCFE1'])
  const totalApptCount = 0 + filteredApptsByPresetCampaigns.length
  return totalApptCount
}



/* ===================== TESTING ========================= */

function run() {
//  var res = GETAGENTPRESETSALES("RICHS", "04/12/2020","04/20/2020")
  var res = GETAGENTSTOTALPRESETS("RICHS", "04/12/2020","04/20/2020")
  Logger.log(res)
}



