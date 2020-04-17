var AprilAppts
Tamotsu.onInitialized(function() {
  AprilAppts = Tamotsu.Table.define({
    sheetName: 'April2020DataImport',
    idColumn: 'apptId',
    test: function() {
      Logger.log('hi')
    }
  },{
    getTotalSold: function(agentName, startingDate, endingDate) {
      const agentNameToUpper = agentName.toUpperCase().split(" " ).join('')
      const allApptDates = AprilAppts.pluck('properAppointmentDate')
      const dateStart = AprilAppts.where(function(appt) {
        appt["properAppointmentDate"] == new Date(startingDate)
      }).all()
      
      const dateEnd = typeof endingDate
      return { startType: dateStart, allappts: allApptDates }
    }
  })
})
