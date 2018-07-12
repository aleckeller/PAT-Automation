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
        # Get incidents and store them in either resolved or open
        api.getIncidents(self)

        # Sort the incidents based on created at
        self.resolved_incidents_array = sorted(set(self.resolved_incidents_array), key=self.resolved_incidents_array.index)
        self.open_incidents_array = sorted(set(self.open_incidents_array), key=self.open_incidents_array.index)

        # Checks for any overlapping incidents. If two incidents overlap, a new incident is created with
        # the lower created at time and the greater end time. This new incident is added to the resolved array
        # and the two old incidents are removed
        api.checkOverlap(self)

        # Need to sort again based on 24 hour time because of the new incident that was added
        self.resolved_incidents_array.sort(key=lambda x: x.twentyfour_created)

        now = datetime.now(pytz.timezone(constants.TIME_ZONE))
        formatted_date = str(now.strftime('%m')) + "-" + str(now.strftime('%d')) + "-" + str(now.strftime('%Y'))
        file = open(formatted_date + ".txt","w")
        outages = False
        if len(self.resolved_incidents_array) > 0:
            outages = True
            index = 1
            file.write("RESOLVED INCIDENTS:\n")
            for incident in self.resolved_incidents_array:
                file.write("    #" + str(index) +  ")  Title: " + str(incident.title) + "\n")
                file.write("         Created At: " + str(incident.created_at) + "\n")
                file.write("         Resolved At: " + str(incident.resolved_time) + "\n")
                file.write("         ID: " + str(incident.id) + "\n")
                file.write("____________________________\n")
                index+=1
        if len(self.open_incidents_array) > 0:
            outages = True
            index = 1
            file.write("STILL OPEN INCIDENTS:")
            for incident in self.open_incidents_array:
                file.write("    #" + str(index) +  ")  Title: " + str(incident.title) + "\n")
                file.write("         Created At: " + str(incident.created_at) + "\n")
                file.write("         ID: " + str(incident.id) + "\n")
                file.write("____________________________\n")
                index+=1
        if not outages:
            file.write("There were no outages on " + formatted_date)

        file.close()
        #excel.loadWorkbook(self)
if __name__ == "__main__":
    main = Main()
