"""
Created on July 2017

@author: Alec Keller
"""
import api
import excel
from datetime import *
import pytz
import constants

class Main:
    def __init__(self):
        self.offset = 0
        self.resolved_incidents_array = []
        self.open_incidents_array = []
        api.getIncidents(self)
        self.resolved_incidents_array = set(self.resolved_incidents_array)
        self.open_incidents_array = set(self.open_incidents_array)
        now = datetime.now(pytz.timezone(constants.TIME_ZONE))
        file = open(str(now.year) + "-" + str(now.month) + "-" + str(now.day) + ".txt","w")
        if len(self.resolved_incidents_array) > 0:
            file.write("RESOLVED INCIDENTS:\n")
            for incident in self.resolved_incidents_array:
                file.write("    " + str(incident.title) + "\n")
                file.write("    " + str(incident.created_at) + "\n")
                file.write("    " + str(incident.resolved_time) + "\n")
                file.write("____________________________\n")
        else:
            file.write("There were no resolved incidents today\n")
        if len(self.open_incidents_array) > 0:
            file.write("STILL OPEN INCIDENTS:")
            for incident in self.open_incidents_array:
                file.write("    " + str(incident.title) + "\n")
                file.write("    " + str(incident.id) + "\n")
                file.write("    " + str(incident.created_at) + "\n")
                file.write("____________________________\n")
        else:
            file.write("There are no incidents that are still open\n")

        file.close()
        #excel.loadWorkbook(self)
if __name__ == "__main__":
    main = Main()
