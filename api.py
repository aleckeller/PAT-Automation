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

def getIncidents(self,incident_id=""):
    now = (datetime.now(pytz.timezone(constants.TIME_ZONE)))
    startOfDay = datetime(now.year, now.month, now.day, tzinfo=now.tzinfo)
    headers = {
        'Accept': 'application/vnd.pagerduty+json;version=2',
        'Authorization': 'Token token={token}'.format(token=API_KEY)
    }

    if incident_id != "":
        url = 'https://api.pagerduty.com/incidents/{id}'.format(id=incident_id)
        payload = {
            'time_zone': constants.TIME_ZONE,
        }
        request = requests.get(url, headers=headers, params=payload)
        return request.json()

    else:
        url = 'https://api.pagerduty.com/incidents'
        #"2018-07-25 23:57:46.651475-04:00"
        payload = {
            'since': startOfDay,
            'until': now,
            'time_zone': constants.TIME_ZONE,
            'limit': constants.LIMIT,
            'sort_by': ["created_at"],
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
                if check == "check-api-umbrella-status-non-prod" or check == "check_docker dtr-" or check == "check_http_alpha.sam.gov" or check == "check_http_beta.sam.gov":
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
    twentyfour_created = created_at.strftime('%H:%M %p')
    #if the incident has been resolved, record end time
    if incident['status'] == 'resolved':
        resolved_time = parser.parse(incident['last_status_change_at'])
        resolved_time_string = resolved_time.strftime('%I:%M %p')
        twentyfour_resolved = resolved_time.strftime('%H:%M %p')
        tdelta = resolved_time - created_at
        # if the incident lasted longer than two minutes, record info
        if tdelta.seconds >= 120:
            incident_obj = ResolvedIncident(check_name,created_at_string,resolved_time_string,incident['id'],twentyfour_created,twentyfour_resolved)
            self.resolved_incidents_array.append(incident_obj)
    else:
        # Get incident number and write to PAT
        incident_obj = OpenIncident(incident['id'],check_name,created_at_string,twentyfour_created)
        self.open_incidents_array.append(incident_obj)

def checkOverlap(self):
    list = self.resolved_incidents_array
    i = 0
    while i < len(list):
        x = i + 1
        while x < len(list):
            if list[i].title == list[x].title:
                if list[i].twentyfour_resolved > list[x].twentyfour_created:
                    #max_resolved_object = null
                    if list[i].twentyfour_resolved > list[x].twentyfour_resolved:
                        max_resolved_object = list[i]
                    else:
                        max_resolved_object = list[x]
                    incident_obj = ResolvedIncident(list[i].title,list[i].created_at,max_resolved_object.resolved_time,"overlap",list[i].twentyfour_created,max_resolved_object.twentyfour_resolved)
                    list[i] = incident_obj
                    list.remove(list[x])
                else:
                    x += 1
            else:
                x += 1
        i += 1
    self.resolved_incidents_array = list
