from datetime import *
import pytz
import requests
import json
import constants
import classes
from classes import *
from dateutil import parser
import sys

API_KEY = sys.argv[1]

def getIncidents(self):
    now = (datetime.now(pytz.timezone(constants.TIME_ZONE)))
    startOfDay = datetime(now.year, now.month, now.day, tzinfo=now.tzinfo)
    url = 'https://api.pagerduty.com/incidents'
    headers = {
        'Accept': 'application/vnd.pagerduty+json;version=2',
        'Authorization': 'Token token={token}'.format(token=API_KEY)
    }
    payload = {
        'since': startOfDay,
        'until': now,
        'time_zone': constants.TIME_ZONE,
        'limit': constants.LIMIT,
        'offset': self.offset
    }

    request = requests.get(url, headers=headers, params=payload)
    parseIncidentsJson(self,request.json())

def parseIncidentsJson(self, json):
    incidents = json['incidents']
    for incident in incidents:
        incident_title = incident['title']
        check_name = ""
        if 'check' in incident['title']:
            split_array = "check" + incident_title.split("check",1)[1]
            check_name = split_array.split(" ")[0].strip()
            for check in constants.HTTPCHECKS:
                if check == "check-api-umbrella-status-non-prod" or check == "check_docker dtr-":
                    if check in check_name:
                        createIncidentObject(self,incident,check_name)
                else:
                    if check == check_name:
                        createIncidentObject(self,incident,check_name)
    if json['more']:
        self.offset += constants.LIMIT
        getIncidents(self)

def createIncidentObject(self,incident,check_name):
    created_at = parser.parse(incident['created_at'])
    created_at_string = created_at.strftime('%I:%M %p')
    #if the incident has been resolved, record end time
    if incident['status'] == 'resolved':
        resolved_time = parser.parse(incident['last_status_change_at'])
        resolved_time_string = resolved_time.strftime('%I:%M %p')
        tdelta = resolved_time - created_at
        # if the incident lasted longer than two minutes, record info
        if tdelta.seconds >= 120:
            incident_obj = ResolvedIncident(check_name,created_at_string,resolved_time_string)
            self.resolved_incidents_array.append(incident_obj)
    else:
        # Get incident number and write to PAT
        incident_obj = OpenIncident(incident['id'],check_name,created_at_string)
        self.open_incidents_array.append(incident_obj)
